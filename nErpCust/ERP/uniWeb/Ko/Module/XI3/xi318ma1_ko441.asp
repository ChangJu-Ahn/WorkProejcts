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
Const BIZ_PGM_ID = "xi318mb1_ko441.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================

Dim C_wc_cd				
Dim C_wc_nm				
Dim C_item_cd			
Dim C_item_nm			
Dim C_spec				
Dim C_lot_no			
Dim C_lot_sub_no		
Dim C_good_on_hand_qty	

'2008-04-15 10:27오후 :: hanc
Dim C_PLANT_CD
Dim C_PRODT_ORDER_NO_S
Dim C_MES_ITEM_CD_S
Dim C_PLANT_CD_TARGET
Dim C_PRODT_ORDER_NO_T
Dim C_MES_ITEM_CD_T
Dim C_JOB_QTY_T
Dim C_CREATE_TYPE
Dim C_SEND_DT
Dim C_ERP_RECEIVE_DT
Dim C_ERP_APPLY_FLAG
Dim C_ERR_DESC

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
Dim iDBSYSDate
Dim LocSvrDate
DIM EndDate

LastDate    = UNIGetLastDay ("<%=BaseDate%>",parent.gServerDateFormat)                                  'Last  day of this month
FirstDate   = UNIGetFirstDay("<%=BaseDate%>",parent.gServerDateFormat)                                  'First day of this month
ExampleDate = UniDateAdd("m", -2, "<%=BaseDate%>",parent.gServerDateFormat)
ExampleDate = UNIConvDateAToB("<%=BaseDate%>" ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
	iDBSYSDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	EndDate = UNIDateAdd("D", 0,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 마지막 날짜 


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
'	C_wc_cd				=  1
'	C_wc_nm				=  2
'	C_item_cd			=  3
'	C_item_nm			=  4
'	C_spec				=  5
'	C_lot_no			=  6
'	C_lot_sub_no		=  7
'	C_good_on_hand_qty	=  8

    C_PLANT_CD            =  1
    C_PRODT_ORDER_NO_S    =  2
    C_MES_ITEM_CD_S       =  3
    C_PLANT_CD_TARGET     =  4
    C_PRODT_ORDER_NO_T    =  5
    C_MES_ITEM_CD_T       =  6
    C_JOB_QTY_T           =  7
    C_CREATE_TYPE         =  8
    C_SEND_DT             =  9
    C_ERP_RECEIVE_DT      =  10
    C_ERP_APPLY_FLAG      =  11
    C_ERR_DESC            =  12


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

	frm1.txtSendStartDt.Text = EndDate
	frm1.txtSendEndDt.Text = EndDate

	frm1.txtPlanStartDt.Text = EndDate
	frm1.txtPlanEndDt.Text = EndDate

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

Sub txtPlanStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanStartDt.Focus
    End If
End Sub

Sub txtPlanEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanEndDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================

Sub txtSendStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSendStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtSendStartDt.Focus
    End If
End Sub

Sub txtSendEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSendEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtSendEndDt.Focus
    End If
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
   
    ggoSpread.Source = frm1.vspdData
                       'Data        Seperator            Column position 
'    ggoSpread.SetCombo "Y"        & vbTab & "N"        , C_StudyOnOffCd
'    ggoSpread.SetCombo "재학" & vbTab & "휴학" , C_StudyOnOffNm
   
    
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
       .MaxCols   = C_ERR_DESC + 1                                                  ' ☜:☜: Add 1 to Maxcols
       Call ggoSpread.ClearSpreadData()
       Call AppendNumberPlace("6","4","2")
       Call GetSpreadColumnPos("A")
				

       ggoSpread.SSSetEdit    C_PLANT_CD            ,"공장(S)" ,8    ,2                  ,     ,100     ,1
       ggoSpread.SSSetEdit    C_PRODT_ORDER_NO_S    ,"제조오더번호(S)" ,13    ,0                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_MES_ITEM_CD_S       ,"품목(S)" ,13    ,0                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_PLANT_CD_TARGET     ,"공장(T)" ,8    ,2                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_PRODT_ORDER_NO_T    ,"제조오더번호(T)" ,13    ,0                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_MES_ITEM_CD_T       ,"품목(T)" ,13    ,0                  ,     ,100     ,2
       ggoSpread.SSSetFloat    C_JOB_QTY_T           ,"지시수량(T)" ,10    , Parent.ggQtyNo   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec ,       ,      ,"P" 
       ggoSpread.SSSetEdit    C_CREATE_TYPE         ,"생성구분" ,8    ,2                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_SEND_DT             ,"ERP수신일시" ,10    ,0                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_ERP_RECEIVE_DT      ,"ERP수신일시" ,13    ,0                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_ERP_APPLY_FLAG      ,"ERP반영여부" ,10    ,2                  ,     ,100     ,2
       ggoSpread.SSSetEdit    C_ERR_DESC            ,"에러내역" ,30    ,0                  ,     ,100     ,2

'       ggoSpread.SSSetEdit    C_wc_cd             ,"작업장코드" ,10    ,0                  ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_wc_nm             ,"작업장명"   ,15    ,0                  ,     ,100     ,1
'       ggoSpread.SSSetEdit    C_item_cd           ,"품목코드"   ,13    ,0                  ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_item_nm           ,"품목명"	    ,20    ,0                  ,     ,100     ,1
'       ggoSpread.SSSetEdit    C_spec			  ,"규격"		,18    ,0                  ,     ,100     ,1
'       ggoSpread.SSSetEdit    C_lot_no            ,"금형번호"   ,15    ,0                  ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_lot_sub_no        ,"파레트번호" ,15    ,0                  ,     ,100     ,2
'       
'       ggoSpread.SSSetFloat   C_good_on_hand_qty  ,"양품수량"   ,15    , Parent.ggQtyNo   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec ,       ,      ,"P" 
       
                             'ColumnPosition     Row
'       ggoSpread.SSSetButton  C_ZipCodePopUp    ,-1
                             'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
'       ggoSpread.SSSetEdit    C_AddressNm       ,"주소"       ,40    ,                   ,     ,     ,2
                             'ColumnPosition     Header            Width  Align(0:L,1:R,2:C)  ComboEditable  Row
'       ggoSpread.SSSetCombo   C_StudyOnOffCd    ,"Y/N"        ,5     ,2                  ,False         ,-1
'       ggoSpread.SSSetCombo   C_StudyOnOffNm    ,"재학/휴학"  ,8     ,2                  ,False         ,-1
                             'Col                Header            Width  Align(0:L,1:R,2:C)  Format         Row
'       ggoSpread.SSSetDate    C_EnrollDT        ,"입학일"     ,15    ,2                  ,parent.gDateFormat   ,-1
'       ggoSpread.SSSetDate    C_GraduatedDT     ,"졸업일"     ,15    ,2                  ,parent.gDateFormat   ,-1
                             'Col                Header            Width  Grp   IntegeralPart              DeciPointpart                                                   Align   Sep    PZ   Min       Max 
'       ggoSpread.SSSetFloat   C_SMoney          ,"용돈"       ,10    ,"6"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec ,       ,      ,"P"
'       ggoSpread.SSSetFloat   C_SMoneyCnt       ,"횟수"       ,15    ,"6"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec ,1      ,True  ,"Z" ,"-12"    ,"10034x30"
'Const ggAmtOfMoneyNo    = "2"      ' Amount No
'Const ggQtyNo           = "3"      ' Quantity No
'Const ggUnitCostNo      = "4"      ' Cost No
'Const ggExchRateNo      = "5"      ' Exchange Rate No

'       call ggoSpread.MakePairsColumn(C_ZipCode,C_ZipCodePopUp)
'       call ggoSpread.MakePairsColumn(C_EnrollDT,C_GraduatedDT,"1")

       Call ggoSpread.SSSetColHidden(C_SEND_DT,C_SEND_DT,True)
       Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

       Call ggoSpread.SSSetSplit2(2)

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
      ggoSpread.SpreadLock       C_PLANT_CD             , -1         , C_ERR_DESC        , -1       
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
                                 'Col          Row         Row2
       ggoSpread.SSSetRequired    C_SID      , pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_SNm      , pvStartRow, pvEndRow
                                 'Col          Row          Row2
       ggoSpread.SSSetProtected   C_AddressNm, pvStartRow, pvEndRow
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
            
'            C_wc_cd				=  iCurColumnPos(1)
'			C_wc_nm				=  iCurColumnPos(2)
'			C_item_cd			=  iCurColumnPos(3)
'			C_item_nm			=  iCurColumnPos(4)
'			C_spec				=  iCurColumnPos(5)
'			C_lot_no			=  iCurColumnPos(6)
'			C_lot_sub_no		=  iCurColumnPos(7)
'			C_good_on_hand_qty	=  iCurColumnPos(8)

            C_PLANT_CD            =  iCurColumnPos(1)  
            C_PRODT_ORDER_NO_S    =  iCurColumnPos(2)  
            C_MES_ITEM_CD_S       =  iCurColumnPos(3)  
            C_PLANT_CD_TARGET     =  iCurColumnPos(4)  
            C_PRODT_ORDER_NO_T    =  iCurColumnPos(5)  
            C_MES_ITEM_CD_T       =  iCurColumnPos(6)  
            C_JOB_QTY_T           =  iCurColumnPos(7)  
            C_CREATE_TYPE         =  iCurColumnPos(8)  
            C_SEND_DT             =  iCurColumnPos(9)  
            C_ERP_RECEIVE_DT      =  iCurColumnPos(10)  
            C_ERP_APPLY_FLAG      =  iCurColumnPos(11)  
            C_ERR_DESC            =  iCurColumnPos(12)  

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
	Call SetToolBar("11000000000011")                                              '☆: Developer must customize

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
   	 Frm1.vspdData.Col = C_StudyOnOffCD :     iDx = Frm1.vspdData.value
     Frm1.vspdData.Col = C_StudyOnOffNM :     Frm1.vspdData.value = iDx
    
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

	Dim strVal
	Dim rdoFlag
	Dim rdoFlagF	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call MakeKeyStream(pDirect)
    With Frm1

        If .rdoFlg1.checked then
    		rdoFlag="%"
    	ElseIf .rdoFlg2.checked Then 
    		rdoFlag ="Y"
    	Else
    		rdoFlag ="N"
    	End If
        
        If .rdoFlgF1.checked then
    		rdoFlagF="%"
    	ElseIf .rdoFlgF2.checked Then 
    		rdoFlagF ="Y"
    	Else
    		rdoFlagF ="N"
    	End If
    	
    	

		strVal = BIZ_PGM_ID & "?txtMode="       & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="  & lgKeyStream         '☜: Query Key
        strVal = strVal     & "&txtPlantCd="    & Trim(.txtPlantCd.value)                       '☜: Query Key
        strVal = strVal     & "&txtItemCds="    & Trim(.txtItemCds.value)                       '☜: Query Key
        strVal = strVal     & "&txtItemCdt="    & Trim(.txtItemCdt.value)                       '☜: Query Key
		strVal = strVal     & "&txtSendStartDt="   & Trim(.txtSendStartDt.Text)
		strVal = strVal     & "&txtSendEndDt="     & Trim(.txtSendEndDt.Text)
		strVal = strVal & "&txtPlanStartDt="   & Trim(.txtPlanStartDt.Text)
		strVal = strVal & "&txtPlanEndDt="     & Trim(.txtPlanEndDt.Text)
		strVal = strVal     & "&rdoFlag="          & Trim(rdoFlag)				'☆: 조회 조건 데이타 
	    strVal = strVal     & "&rdoFlagF="          & Trim(rdoFlagF)				'☆: 조회 조건 데이타 
        strVal = strVal     & "&txtMaxRows="    & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey        '☜: Next key tag
    End With
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
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

       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
       .txtMaxRows.value     = lGrpCnt - 1
       .txtSpread.value      = strDel & strVal


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
	Call SetToolbar("11000000001111")                                              '☆: Developer must customize
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
			Case C_ZipCodePopUp
				.Col = Col - 1
				.Row = Row
				Call OpenZipCode(.Text,Row)
			End Select
		End If    
	End With
End Sub


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
         Case  C_StudyOnOffNm
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_StudyOnOffCd
                Frm1.vspdData.value = iDx
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

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfos(Byval strCode)
    
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode			' Item Code
	arrParam(2) = ""				' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field명(0)
	arrField(1) = 2 '"ITEM_NM"			' Field명(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfos(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCds.focus

End Function
Function SetItemInfos(byval arrRet)
	frm1.txtItemCds.Value    = arrRet(0)		
	frm1.txtItemNms.Value    = arrRet(1)		
End Function

Function OpenItemInfot(Byval strCode)
    
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode			' Item Code
	arrParam(2) = ""				' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field명(0)
	arrField(1) = 2 '"ITEM_NM"			' Field명(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfot(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCdt.focus

End Function
Function SetItemInfot(byval arrRet)
	frm1.txtItemCdt.Value    = arrRet(0)		
	frm1.txtItemNmt.Value    = arrRet(1)		
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Device Change 수신</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>MES송신기간</TD>
								    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtSendStartDt 	CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtSendEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
    								</TD>	
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목(S)</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCds" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfos frm1.txtItemCds.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNms" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목(T)</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCdt" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfot frm1.txtItemCdt.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNmt" SIZE=25 tag="14"></TD>
								</TR>								
									<TR>
									<TD CLASS=TD5 NOWRAP>정상처리여부</TD>
									<TD CLASS=TD6 NOWRAP>
									    <INPUT TYPE="RADIO" NAME="rdoFlgF" ID="rdoFlgF1" CLASS="RADIO" tag="11" Value="A" CHECKED><LABEL FOR="rdoFlgF1">전체</LABEL>
									    <INPUT TYPE="RADIO" NAME="rdoFlgF" ID="rdoFlgF2" CLASS="RADIO" tag="11" Value="Y" ><LABEL FOR="rdoFlgF2">정상처리</LABEL>
				     				    <INPUT TYPE="RADIO" NAME="rdoFlgF" ID="rdoFlgF3" CLASS="RADIO" tag="11" Value="N" ><LABEL FOR="rdoFlgF3">처리중오류</LABEL></TD>
									</TD>								
									<TD CLASS=TD5 NOWRAP>생산실적기간</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtPlanStartDt	CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtPlanEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									<TD CLASS=TD5 NOWRAP><div style="display:none">ERP반영여부</div></TD>
									<TD CLASS=TD6 NOWRAP><div style="display:none">
									    <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg1" CLASS="RADIO" tag="11" Value="A" CHECKED><LABEL FOR="rdoFlg1">전체</LABEL>
				     				    <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg2" CLASS="RADIO" tag="11" Value="Y" ><LABEL FOR="rdoFlg2">반영</LABEL>
				     				    <INPUT TYPE="RADIO" NAME="rdoFlg" ID="rdoFlg3" CLASS="RADIO" tag="11" Value="N" ><LABEL FOR="rdoFlg3">미반영</LABEL></div></TD>
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
		<TD  HEIGHT=3></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR style="display:none">
					<TD>&nbsp;</TD>
					<TD><!--<BUTTON NAME="btnRec" CLASS="CLSMBTN" ONCLICK="vbscript:RecMes()" >MES정보수신</BUTTON> -->
					&nbsp;<BUTTON NAME="btnReg" CLASS="CLSMBTN" ONCLICK="vbscript:RegProd()" >D/C 제조오더등록</BUTTON>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					&nbsp;<BUTTON NAME="btnReg" CLASS="CLSMBTN" ONCLICK="vbscript:RegProd1()" >처리중오류DATA등록</BUTTON></TD>				
					<TD>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtPlantCd"     TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

