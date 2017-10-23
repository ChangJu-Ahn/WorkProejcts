<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: 인산/급여관리 
*  2. Function Name        	: 소급분관리 
*  3. Program ID           	: H8005ma1
*  4. Program Name         	: 소급급여부서별 조회 
*  5. Program Desc         	: multi Sample
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/04/18
*  8. Modified date(Last)  	: 2003/06/13
*  9. Modifier (First)     	: Hwang Jeong Won
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment             	:
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "h8005mb1.asp"                                      'Biz Logic ASP 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop   
Dim lgFrQueryFlg, lgToQueryFlg

Dim C_DEPT_CD
Dim C_EMP_CNT_AMT
Dim C_COMPUTE_0003_AMT
Dim C_COMPUTE_0004_AMT
Dim C_COMPUTE_0005_AMT
Dim C_COMPUTE_0006_AMT
Dim C_RETRO1_AMT
Dim C_RETRO2_AMT

Dim C_DEPT_CD2
Dim C_EMP_CNT_AMT2
Dim C_COMPUTE_0003_AMT2
Dim C_COMPUTE_0004_AMT2
Dim C_COMPUTE_0005_AMT2
Dim C_COMPUTE_0006_AMT2
Dim C_RETRO1_AMT2
Dim C_RETRO2_AMT2
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
 
    C_DEPT_CD = 1     	
    C_EMP_CNT_AMT = 2    
    C_COMPUTE_0003_AMT = 3     
    C_COMPUTE_0004_AMT = 4   
    C_COMPUTE_0005_AMT = 5     
    C_COMPUTE_0006_AMT = 6     
    C_RETRO1_AMT = 7 
    C_RETRO2_AMT = 8
    
    C_DEPT_CD2 = 1     	
    C_EMP_CNT_AMT2 = 2    
    C_COMPUTE_0003_AMT2 = 3     
    C_COMPUTE_0004_AMT2 = 4   
    C_COMPUTE_0005_AMT2 = 5     
    C_COMPUTE_0006_AMT2 = 6     
    C_RETRO1_AMT2 = 7 
    C_RETRO2_AMT2 = 8
    
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
    Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtpay_yymm_dt.Focus()
		
	frm1.txtpay_yymm_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtpay_yymm_dt.Month = strMonth 
    lgFrQueryFlg = false	
    lgToQueryFlg = false	
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
	Dim strYear, strMonth
   	Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")
    
   	strYear = Trim(frm1.txtpay_yymm_dt.year)
   	strMonth = Trim(frm1.txtpay_yymm_dt.month)
   	If Len(strMonth) = 1 then
   		strMonth = "0" & strMonth
   	end If 
   	
    lgKeyStream  = strYear & strMonth & Parent.gColSep       'You Must append one character(Parent.gColSep)
    lgKeyStream  = lgKeyStream & frm1.txtFr_internal_cd.Value & Parent.gColSep
    lgKeyStream  = lgKeyStream & frm1.txtTo_internal_cd.Value & Parent.gColSep
    lgKeyStream  = lgKeyStream & StrDt & Parent.gColSep  
End Sub 

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()

	Dim intIndex 
	lgFrQueryFlg = false	
	lgToQueryFlg = false	

	With frm1.vspdData
	    
	    ggoSpread.Source = frm1.vspdData2
        ggoSpread.UpdateRow 1
        
        frm1.vspdData2.Col = 0
        frm1.vspdData2.Text = "합계"
        frm1.vspdData2.Col = C_DEPT_CD2        
        frm1.vspdData2.Text = frm1.vspdData.MaxRows
        frm1.vspdData2.Col = C_EMP_CNT_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_EMP_CNT_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_COMPUTE_0003_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_COMPUTE_0003_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_COMPUTE_0004_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_COMPUTE_0004_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_COMPUTE_0005_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_COMPUTE_0005_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_COMPUTE_0006_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_COMPUTE_0006_AMT, 1, .MaxRows, FALSE, -1, -1, "V")
        frm1.vspdData2.Col = C_RETRO1_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_RETRO1_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_RETRO2_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_RETRO2_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
   End With
       
   Call SetSpreadLock("B")
   
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call initSpreadPosVariables()   'sbk 

    If pvSpdNo = "" OR pvSpdNo = "A" Then

	    With frm1.vspdData	

            ggoSpread.Source = frm1.vspdData
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
           
           .MaxCols = C_RETRO2_AMT + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	       .Col = .MaxCols															'공통콘트롤 사용 Hidden Column
           .ColHidden = True
    
           .MaxRows = 0
           ggoSpread.ClearSpreadData

           Call GetSpreadColumnPos("A") 'sbk
	
           Call AppendNumberPlace("6","15","0")

	    		ggoSpread.SSSetEdit C_DEPT_CD, "부서", 10
	    		ggoSpread.SSSetFloat C_EMP_CNT_AMT, "인원", 8,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
	    		ggoSpread.SSSetFloat C_COMPUTE_0003_AMT, "원지급분기본수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    		ggoSpread.SSSetFloat C_COMPUTE_0004_AMT, "원지급분기타수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
	    		ggoSpread.SSSetFloat C_COMPUTE_0005_AMT, "인상분기본수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    		ggoSpread.SSSetFloat C_COMPUTE_0006_AMT, "인상분기타수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	
	    		ggoSpread.SSSetFloat C_RETRO1_AMT, "소급분기본수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    		ggoSpread.SSSetFloat C_RETRO2_AMT, "소급분기타수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	
	       .ReDraw = true

           Call SetSpreadLock("A")
    
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
    
        With frm1.vspdData2

            ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk
           
	       .ReDraw = false

           .MaxCols = C_RETRO2_AMT2 + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
	       .Col = .MaxCols															'공통콘트롤 사용 Hidden Column
           .ColHidden = True
    
           .MaxRows = 0
           ggoSpread.ClearSpreadData

           .DisplayColHeaders = False

           Call GetSpreadColumnPos("B") 'sbk
	
           Call AppendNumberPlace("6","15","0")
	    		
	    	    ggoSpread.SSSetFloat C_DEPT_CD2, "부서", 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    		ggoSpread.SSSetFloat C_EMP_CNT_AMT2, "인원", 8,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
	    		ggoSpread.SSSetFloat C_COMPUTE_0003_AMT2, "원지급분기본수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    		ggoSpread.SSSetFloat C_COMPUTE_0004_AMT2, "원지급분기타수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    
	    		ggoSpread.SSSetFloat C_COMPUTE_0005_AMT2, "인상분기본수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    		ggoSpread.SSSetFloat C_COMPUTE_0006_AMT2, "인상분기타수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	
	    		ggoSpread.SSSetFloat C_RETRO1_AMT2, "소급분기본수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	    		ggoSpread.SSSetFloat C_RETRO2_AMT2, "소급분기타수당", 16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	
	       .ReDraw = true
	
           Call SetSpreadLock("B")
    
        End With
    End If
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
        If pvSpdNo = "A" Then
            ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()    

        ElseIf pvSpdNo = "B" Then
             ggoSpread.Source = frm1.vspdData2
            .vspdData2.ReDraw = False
              ggoSpread.SpreadLock C_DEPT_CD2, -1 , C_DEPT_CD2, -1            
              ggoSpread.SpreadLock C_EMP_CNT_AMT2, -1 , C_EMP_CNT_AMT2, -1
              ggoSpread.SpreadLock C_COMPUTE_0003_AMT2, -1 , C_COMPUTE_0003_AMT2, -1
              ggoSpread.SpreadLock C_COMPUTE_0004_AMT2, -1 , C_COMPUTE_0004_AMT2, -1
              ggoSpread.SpreadLock C_COMPUTE_0005_AMT2, -1 , C_COMPUTE_0005_AMT2, -1
	          ggoSpread.SpreadLock C_COMPUTE_0006_AMT2, -1 , C_COMPUTE_0006_AMT2, -1
	          ggoSpread.SpreadLock C_RETRO1_AMT2, -1 , C_RETRO1_AMT2, -1
	          ggoSpread.SpreadLock C_RETRO2_AMT2, -1 , C_RETRO2_AMT2, -1
	          ggoSpread.SSSetProtected   .vspdData2.MaxCols   , -1, -1
            .vspdData2.ReDraw = True
        End If

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False      
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
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

            C_DEPT_CD = iCurColumnPos(1)
            C_EMP_CNT_AMT = iCurColumnPos(2)
            C_COMPUTE_0003_AMT = iCurColumnPos(3)
            C_COMPUTE_0004_AMT = iCurColumnPos(4)
            C_COMPUTE_0005_AMT = iCurColumnPos(5)   
            C_COMPUTE_0006_AMT = iCurColumnPos(6)
            C_RETRO1_AMT = iCurColumnPos(7)
            C_RETRO2_AMT = iCurColumnPos(8)

      Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_DEPT_CD2 = iCurColumnPos(1)     	
            C_EMP_CNT_AMT2 = iCurColumnPos(2)    
            C_COMPUTE_0003_AMT2 = iCurColumnPos(3)     
            C_COMPUTE_0004_AMT2 = iCurColumnPos(4)   
            C_COMPUTE_0005_AMT2 = iCurColumnPos(5)     
            C_COMPUTE_0006_AMT2 = iCurColumnPos(6)     
            C_RETRO1_AMT2 = iCurColumnPos(7) 
            C_RETRO2_AMT2 = iCurColumnPos(8)
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
 
    Call InitSpreadSheet("")                                                          'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables

    Call SetDefaultVal

    Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, Parent.gDateFormat, 2) '<==== 싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.    

    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어    
    Call CookiePage (0)                                                             '☜: Check Cookie

	
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim CFlag,RType
    
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
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	Call InitVariables                                                           '⊙: Initializes local global variables    
   
    lgFrQueryFlg = true	
	    
  	call  txtFr_dept_cd_onchange() 
	if lgFrQueryFlg = false then
  		Exit Function
  	end if
       
    
    lgToQueryFlg = true
    call txtTo_dept_cd_onchange() 
  	if lgToQueryFlg = false then
  		Exit Function
  	end if
    
    	
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept ,IntRetCd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    
    If fr_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,StrDt, rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,StrDt, rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd= "") AND (To_dept_cd="") Then       
    Else
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtfr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
        
    END IF       
    Call MakeKeyStream("X")
   	Call DisableToolBar(Parent.TBC_QUERY)
    
	IF DBQUERY =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
   	Call DisableToolBar(Parent.TBC_SAVE)
		IF DBSAVE =  False Then
			Call RestoreToolBar()
			Exit Function
		End If
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	With Frm1.VspdData
           .Col  = C_MAJORCD
           .Row  = .ActiveRow
           .Text = ""
    End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow
       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
	
	ggoSpread.Source = frm1.vspdData2 
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
    
    If isEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
	Elseif	UCase(gActiveSpdSheet.id) = "VASPREAD" Then
		ggoSpread.Source = frm1.vspdData2 
		Call ggoSpread.SaveSpreadColumnInf()
	End if

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	dim temp
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")      
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ReOrderingSpreadData()

	temp = GetSpreadText(frm1.vspdData,1,1,"X","X")

	if temp <>"" then
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.InsertRow
		Call InitData()
	end if
End Sub

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	If   LayerShowHide(1) = False Then
     		Exit Function
	End If
	
    Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
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
	
    DbSave = False                                                          
    
    If   LayerShowHide(1) = False Then
     	Exit Function
    End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                    .vspdData.Col = C_MajorCd	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_MajorNm	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_MinorLen	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_TypeCd    : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                   .vspdData.Col = C_MajorCd	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_MajorNm	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_MinorLen	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_TypeCd     : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & Parent.gColSep
                                                  strDel = strDel & lRow & Parent.gColSep
                   .vspdData.Col = C_MajorCd    : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep									
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = Parent.UID_M0002
       .txtUpdtUserId.value  = Parent.gUsrID
       .txtInsrtUserId.value = Parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
   	Call DisableToolBar(Parent.TBC_DELETE)
	IF DBDELETE =  False Then
		Call RestoreToolBar()
		Exit Function
	End If														'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.InsertRow
    Call SetSpreadLock("B")    
    Call InitData()
	Call SetToolbar("1100000000011111")									
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
   	Call DisableToolBar(TBC_QURERY)
		IF DBQUERY =  False Then
			Call RestoreToolBar()
			Exit Function
		End If
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)	
	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True	  
	
	If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If  
		
	strBasDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
    arrParam(1) = strBasDt
    arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
			   frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
        End Select	
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtFr_dept_cd.value = arrRet(0)
               .txtFr_dept_nm.value = arrRet(1)
               .txtFr_internal_cd.value = arrRet(2)               
			   .txtFr_dept_cd.focus
             Case "1"  
               .txtTo_dept_cd.value = arrRet(0)
               .txtTo_dept_nm.value = arrRet(1) 
               .txtTo_internal_cd.value = arrRet(2)               
               .txtTo_dept_cd.focus
        End Select
	End With
End Function       		
'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
sub txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")
	
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""   
	
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , StrDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
			if lgFrQueryFlg = true	then
				If  IntRetCd = -1  then
    				Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
				Else
					Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
				End if
			end if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus()
           
            lgFrQueryFlg = false
            Exit SUB
            
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
	
        End if 
    End if  
    
End sub
'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Sub txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim StrDt
    StrDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtpay_yymm_dt.Year, Right("0" & frm1.txtpay_yymm_dt.month , 2), "01")
	
	
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
	
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , StrDt , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	    
			if lgToQueryFlg = true then
				If  IntRetCd = -1 then
    				Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
				Else
					Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
				End if
			end if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus()
   
            lgToQueryFlg = false
                
            Exit Sub
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
	    End if 
    End if  
    
End sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_TYPENm
                iDx = Frm1.vspdData.value
   	            Frm1.vspdData.Col = C_TYPECd
                Frm1.vspdData.value = iDx
         Case Else
    End Select    
             
   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000101111")

    gMouseClickStatus = "SPC" 

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData2
   
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
End Sub    

Sub vspdData2_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SP1C" Then
          gMouseClickStatus = "SP1CR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData.Col = pvCol1
    frm1.vspdData2.ColWidth(pvCol1) = frm1.vspdData.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData2.Col = pvCol1
    frm1.vspdData.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col , ByVal Row, ByVal newCol , ByVal newRow ,Cancel )
    frm1.vspdData2.Col = newCol
    frm1.vspdData2.Action = 0
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData2.LeftCol=NewLeft   	
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData.LeftCol=NewLeft   	
		Exit Sub
	End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
'==========================================================================================
'   Event Name : txtpay_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtFr_dept_cd_OnKeydown()
	dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
		
	End Select		
	
End Sub

Sub txtTo_dept_cd_OnKeydown()
	dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
		
	End Select		
	
End Sub

'=======================================================================================================
'   Event Name : txtDilig_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtpay_yymm_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub
'=======================================================================================================
'   Event Name : txtBas_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtpay_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtpay_yymm_dt.Action = 7
        frm1.txtpay_yymm_dt.focus
    End If
     lgBlnFlgChgValue = True
End Sub


Sub txtpay_yymm_dt_Change()
    lgBlnFlgChgValue = True
End Sub
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL="no" TABINDEX="-1">
<FORM NAME=Frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>소급급여부서별조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20 width=100%>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
							    <TD CLASS=TD5 NOWRAP>조회년월</TD>
			                    <TD CLASS="TD6" NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> id=txtpay_yymm_dt NAME="txtpay_yymm_dt" CLASS=FPDTYYYYMM  title=FPDATETIME ALT="조회년월" tag="12X1" VIEWASTEXT> </OBJECT></TD>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP>
								        <INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                            <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40" tag="14XXXU">
		                                <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">&nbsp;~&nbsp;
		                                <INPUT NAME="txtto_dept_cd" ALT="부서코드" TYPE="Text" SIZE="10" MAXLENGTH="10" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							            <INPUT NAME="txtto_dept_nm" ALT="부서코드명" TYPE="Text"SIZE="20" MAXLENGTH="40" tag="14XXXU">
							            <INPUT NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU"></TD>
							</TR>							
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%> >
							<TR>
								<TD HEIGHT="100%" width=100%>
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread>
										<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
								</TD>
							</TR>							
							
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=44 WIDTH=100% valign=Top>
                    		<TABLE <%=LR_SPACE_TYPE_60%> >
	                	    	<TR>
						            <TD HEIGHT="100%" width=100%>
							            <OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1>
								           <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
							            </OBJECT>
						            </TD>
					            </TR>
					    </TABLE>
                    	</TD>
					</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>		
		<TD HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


