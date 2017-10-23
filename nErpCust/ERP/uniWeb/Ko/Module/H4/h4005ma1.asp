<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: Multi Sample
*  3. Program ID           	: H4005ma1
*  4. Program Name         	: H4005ma1
*  5. Program Desc         	: 일일근태등록 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/28
*  8. Modified date(Last)  	: 2003/06/11
*  9. Modifier (First)     	: Hwang Jeong Won
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment              	: 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "h4005mb1.asp"                                      '비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd

Dim C_EMPNO
Dim C_EmpPopup 
Dim C_EMPNM
Dim C_DEPT
Dim C_DEPT_NM
Dim C_GRD 
Dim C_CD
Dim C_CdPopup
Dim C_NM
Dim C_HOUR
Dim C_MIN
Dim C_FLAG
Dim IsOpenPop          

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================

Sub initSpreadPosVariables()  
	 C_EMPNO = 1
	C_EmpPopup = 2
	C_EMPNM = 3
	C_DEPT = 4
	C_DEPT_NM =5
	C_GRD = 6
	C_CD = 7
	C_CdPopup = 8
	C_NM = 9
	C_HOUR = 10
	C_MIN = 11
	C_FLAG = 12
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
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
	
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtValidDt.focus

	frm1.txtValidDt.Year = strYear 		 '년월일 default value setting
	frm1.txtValidDt.Month = strMonth 
	frm1.txtValidDt.Day = strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
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
	Dim strFrDept, strToDept,IntRetCd
   
    Call txtCd_OnChange()

    lgKeyStream   = Frm1.txtValidDt.Text & parent.gColSep 
    lgKeyStream   = lgKeyStream & Frm1.txtCd.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtFr_internal_cd.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtTo_internal_cd.Value & parent.gColSep
    
    lgKeyStream   = lgKeyStream & Frm1.txtEmpNo.Value & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim strFlag
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_FLAG
			strFlag = .Value
			
			If strFlag = "1" or strFlag = "3" Then
				ggoSpread.SpreadLock    C_HOUR, intRow, C_MIN, intRow
			End If							
		Next	
	End With	
End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	    .ReDraw = false
        .MaxCols = C_FLAG + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0
		Call GetSpreadColumnPos("A")  
	
	    Call AppendNumberPlace("6","2","0")

        ggoSpread.SSSetEdit  C_EMPNO            , "사번", 10,,, 13, 2
        ggoSpread.SSSetButton C_EmpPopup
        ggoSpread.SSSetEdit  C_EMPNM            , "성명", 16
        ggoSpread.SSSetEdit  C_DEPT             , "부서코드", 10
        ggoSpread.SSSetEdit  C_DEPT_NM          , "부서", 24
        ggoSpread.SSSetEdit  C_GRD              , "직위", 16
        ggoSpread.SSSetEdit  C_CD               , "근태코드", 10,,, 2
        ggoSpread.SSSetButton C_CdPopUp
        ggoSpread.SSSetEdit  C_NM            , "근태명", 20
        ggoSpread.SSSetFloat C_HOUR,"시간" ,8,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","23"
        ggoSpread.SSSetFloat C_MIN,"분" ,8,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","59"
		ggoSpread.SSSetEdit  C_FLAG             , "flag", 10

        Call ggoSpread.MakePairsColumn(C_EMPNO,C_EmpPopup)
        Call ggoSpread.MakePairsColumn(C_CD,C_CdPopUp)
        
        call ggoSpread.SSSetColHidden(C_DEPT,C_DEPT,True)
        call ggoSpread.SSSetColHidden(C_FLAG,C_FLAG,True)
		        
	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
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
            
			C_EMPNO			= iCurColumnPos(1)
			C_EmpPopup		= iCurColumnPos(2)
			C_EMPNM			= iCurColumnPos(3)
			C_DEPT			= iCurColumnPos(4)
			C_DEPT_NM		= iCurColumnPos(5)
			C_GRD			= iCurColumnPos(6)
			C_CD			= iCurColumnPos(7)	
			C_CdPopup		= iCurColumnPos(8) 
			C_NM			= iCurColumnPos(9)
			C_HOUR			= iCurColumnPos(10)
			C_MIN			= iCurColumnPos(11)
			C_FLAG			= iCurColumnPos(12)
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    C_EMPNO, -1, C_EMPNO
        ggoSpread.SpreadLock    C_EmpPopup, -1, C_EmpPopup        
        ggoSpread.SpreadLock    C_EMPNM, -1, C_EMPNM
        ggoSpread.SpreadLock    C_DEPT, -1, C_DEPT
        ggoSpread.SpreadLock    C_DEPT_NM, -1, C_DEPT_NM
        ggoSpread.SpreadLock    C_GRD, -1, C_GRD        
        ggoSpread.SpreadLock    C_CD, -1, C_CD
        ggoSpread.SpreadLock    C_CdPopup, -1, C_CdPopup        
        ggoSpread.SpreadLock    C_NM, -1, C_NM
        ggoSpread.SSSetRequired	C_HOUR, -1, -1
        ggoSpread.SSSetRequired	C_MIN, -1, -1
'        ggoSpread.SpreadLock    C_HOUR, -1, C_HOUR
 '       ggoSpread.SpreadLock    C_MIN, -1, C_MIN
   	    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1 
        .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
       .vspdData.ReDraw = False
         ggoSpread.SSSetRequired		C_EMPNO, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_EMPNM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_DEPT, pvStartRow, pvEndRow   '히든이라도 해당 처리를 해줄것(안그러면 엑셀의 자료를 붙여넣기 시 처리가 안됨)
         ggoSpread.SSSetProtected		C_DEPT_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_GRD, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_CD, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_NM, pvStartRow, pvEndRow
         
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
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
	Call ggoOper.FormatDate(frm1.txtValidDt, parent.gDateFormat, 1)

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
       
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
    Dim IntRetCD 
    Dim strFrDept, strToDept
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If   

    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept
    
    Fr_dept_cd = frm1.txtFrDept.value
    To_dept_cd = frm1.txtToDept.value
   
    If txtCd_OnChange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Function
    End if
 
    If txtFrDept_Onchange() Then
        Exit Function
    End if    
    If txtToDept_Onchange() Then
        Exit Function
    End if    
    If txtEmpNo_Onchange() Then
        Exit Function
    End if  
    If fr_dept_cd = "" then    
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept				
		frm1.txtFrDeptNm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtToDeptNm.value = ""
	End If  
    
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_internal_cd.value = ""
            frm1.txtTo_internal_cd.value = ""
            frm1.txtFrDept.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF 
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                        '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    Call DisableToolBar(parent.TBC_QUERY)
	If DbQuery = False Then
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
    
    FncDelete = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD
    Dim strReturn_value, strSQL
    Dim HFlag,MFlag,Rowcnt
    Dim strVdate
    Dim strWhere
    Dim strDay_time
    
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
    HFlag = False      '올바른 값입력 
    MFlag = False    

    For Rowcnt = 1 To frm1.vspdData.MaxRows
        frm1.vspdData.Row = Rowcnt
        frm1.vspdData.Col = 0

        Select Case frm1.vspdData.Text
           
            Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
				frm1.vspdData.Col = C_EMPNM
				If IsNull(Trim(frm1.vspdData.Text)) OR Trim(frm1.vspdData.Text) = "" Then
				    Call DisplayMsgBox("800048","X","X","X")
					Exit Function
				end if            
                frm1.vspdData.Col = C_NM
                If IsNull(Trim(frm1.vspdData.Text)) OR Trim(frm1.vspdData.Text) = "" Then
                    Call DisplayMsgBox("800099","X","X","X")
                    frm1.vspdData.Action = 0
                    Set gActiveElement = document.activeElement
                    Exit Function
                End if
				strVDate = UNIConvDate(frm1.txtValidDt.text)
				frm1.vspdData.col = C_EMPNO
				strWhere = " EMP_NO=" & FilterVar(frm1.vspdData.Text, "''", "S") 
				strWhere = strWhere & " AND isnull(retire_dt," & FilterVar(strVDate, "''", "S") & ") < " & FilterVar(strVDate, "''", "S")

				intRetCD = CommonQueryRs(" Retire_dt " ," haa010t a",strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   
				
				If intRetCD= true Then  '퇴직자의 경우 퇴직일 다음날 부터 입력 불가 
					frm1.vspdData.col = C_EMPNM
					Call DisplayMsgBox("800494","X",frm1.vspdData.text,UNIDateClientFormat(Trim(Replace(lgF0,Chr(11),""))))
					Set gActiveElement = document.activeElement
					exit function
				End if
		    
				frm1.vspdData.Col = C_CD            
				IntRetCD = CommonQueryRs(" DAY_TIME "," hca010t a"," DILIG_CD= " & FilterVar(frm1.vspdData.Text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   
                
                strDay_time = Trim(Replace(lgF0,Chr(11),""))
                
				If strDay_time = "2" Then  '근태코드테이블(hca010t)의 일수구분이 시간일 경우만 해당됨.
					frm1.vspdData.Row = Rowcnt
					frm1.vspdData.Col = C_HOUR            

					If frm1.vspdData.Text = "0" Then
						HFlag = True     '0을 입력                
					Else
						HFlag = False
					End If

					frm1.vspdData.Col = C_MIN	         

					If frm1.vspdData.Text = "0" Then
						MFlag = True
					Else
						MFlag = False
					End If            

					If HFlag And MFlag Then
						Call DisplayMsgBox("800443","x","시간/분","0")                           '시간은 0 보다 커야합니다.
						frm1.vspdData.Col = C_HOUR
						frm1.vspdData.Action = 0
						Set gActiveElement = document.ActiveElement   
						Exit Function
					End If       
				End If
	    End Select
    Next

	strReturn_value = "N"
    strSQL = " org_cd = " & FilterVar("1", "''", "S") & "  AND pay_gubun = " & FilterVar("Z", "''", "S") & "  AND PAY_TYPE = " & FilterVar("#", "''", "S") & " "
    IntRetCD = CommonQueryRs(" close_type, convert(char(10),close_dt,20), emp_no "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If  IntRetCd = false then
        strReturn_value = "Y"
    else
		Dim strCloseDt, strValidDt
		 
		strCloseDt = UniConvDateToYYYYMMDD(Replace(lgF1, Chr(11), ""),parent.gServerDateFormat,"")
		strValidDt = UniConvDateToYYYYMMDD(frm1.txtValidDt.text,parent.gDateFormat,"")
	
        Select Case Replace(lgF0, Chr(11), "")
            Case "1"    '마감형태 : 정상 
                if strCloseDt <= strValidDt then
                    strReturn_value = "Y"
                else
                    strReturn_value = "N"
                end if

            Case "2"    '마감형태 : 마감 
                if strCloseDt < strValidDt then
                    strReturn_value = "Y"
                else
                    strReturn_value = "N"
                end if
                
        end Select
    end if
    if  strReturn_value = "N" then
        Call DisplayMsgBox("800291","X","X","X")
        exit function
    end if
    
    FncSave = True                                            
    
	Call DisableToolBar(parent.TBC_SAVE)
	If DbSave = False Then                                    '☜: Save db data     Processing is OK
		Call RestoreToolBar()
        Exit Function
    End If
    
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    FncCopy = False           
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
            .Col = C_EMPNO
		    .Focus
		    .Action = 0 ' go to 
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	   .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
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
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function
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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	 If LayerShowHide(1) = False then
    		Exit Function 
    	End if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
	
    DbSave = False                                                          
    
     If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
        
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  strVal = strVal & Trim(Frm1.txtValidDt.Text) & parent.gColSep
                                        
                    .vspdData.Col = C_EMPNO		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CD		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HOUR			    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_MIN		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_NM		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT		        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
												  strVal = strVal & Trim(Frm1.txtValidDt.Text) & parent.gColSep
												  
                    .vspdData.Col = C_EMPNO		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CD		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HOUR			    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_MIN		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_NM		        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                                  strDel = strDel & Trim(Frm1.txtValidDt.Text) & parent.gColSep
                                                 
                    .vspdData.Col = C_EMPNO		        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CD		        : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtMode.value        = parent.UID_M0002
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

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If    
    
    Call DisableToolBar(parent.TBC_DELETE)
	If DbDelete = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncDelete = True                                                        '⊙: Processing is OK

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")									
	frm1.vspdData.focus	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
    Call InitVariables															'⊙: Initializes local global variables
	Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function
'----------------------------------------  OpenEmp()  ------------------------------------------
'	Name : OpenEmp()
'	Description : Employee PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = UCase(Trim(frm1.txtEmpNo.value))			<%' Code Condition%>
		arrParam(1) = ""'frm1.txtEmpNm.value		    ' Name Cindition
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			<%' Code Condition%>
		
	End If
	
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtEmpNo.focus
		Else 'spread
			frm1.vspdData.Col = C_EMPNO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetEmp()  ------------------------------------------------
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetEmp(Byval arrRet, Byval iWhere)
		
	Dim strFg, strValidDt
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmpNo.value = arrRet(0)
			.txtEmpNm.value = arrRet(1)
			.txtEmpNo.focus
		Else 'spread
			.vspdData.Col = C_EMPNO
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_EMPNM
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_GRD
			.vspdData.Text = arrRet(3)
			
' 부서정보 가져오는 로직 정리 2003.9.17 by lsn 	
			call SetSpreadDept(arrRet(0))
   	                        
			.vspdData.Col = C_EMPNO
			.vspdData.action =0
		End If
	End With
End Function
'-------------------------------------------------------------------------------------------------------
'	Name : SetSpreadDept()
'	Description : spread에서 부서코드,부서명 setting - 2003.9.17 by lsn 	
'---------------------------------------------------------------------------------------------------------
Sub SetSpreadDept(Byval emp_no)
	Dim strFg,strValidDt	
	
	strValidDt = UNIConvDate(frm1.txtValidDt.Text)
    Call CommonQueryRs("dbo.ufn_H_get_dept_cd( " & FilterVar(emp_no, "''", "S") & ", " & FilterVar(strValidDt, "''", "S") & ")","", "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strFg = Replace(lgF0, Chr(11), "")

	Frm1.vspdData.Col = C_DEPT
   	Frm1.vspdData.value = strFg

    Call CommonQueryRs("dbo.ufn_GetDeptName( " & FilterVar(strFg, "''", "S") & ", " & FilterVar(strValidDt, "''", "S") & ")","", "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strFg = Replace(lgF0, Chr(11), "")

	Frm1.vspdData.Col = C_DEPT_NM
   	Frm1.vspdData.value = strFg
End Sub
'----------------------------------------  OpenDept()  ------------------------------------------
'	Name : OpenDept()
'	Description : Dept PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = UCase(Trim(frm1.txtfrDept.value))			' from 조건부에서 누른 경우 Code Condition
	Else 
		arrParam(0) = UCase(Trim(frm1.txttoDept.value))			' to 조건부에서 누른 경우 Code Condition
	End If
	
	arrParam(1) = Frm1.txtValidDt.Text
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 
			frm1.txtfrDept.focus
		Else 
			frm1.txttoDept.focus
		End If	
		Exit Function
	Else
		If iWhere = 0 Then 
			frm1.txtfrDept.value = arrRet(0)
			frm1.txtfrDeptNm.value = arrRet(1)
			frm1.txtfrDept.focus
		Else 
			frm1.txttoDept.value = arrRet(0)
			frm1.txttoDeptNm.value = arrRet(1)
			frm1.txttoDept.focus
		End If	
	End If	
			
End Function

'------------------------------------------  OpenAttend()  -------------------------------------------
'	Name : OpenAttend()
'	Description : 근태코드 PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenAttend(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "근태코드 팝업"			' 팝업 명칭 
	arrParam(1) = "hca010t"						' TABLE 명칭 
    
    If iWhere = 0 Then 'TextBox(Condition)
		arrParam(2) = UCase(Trim(frm1.txtCd.value))				' Code Condition
	Else 'spread
		frm1.vspdData.Col = C_CD
	    arrParam(2) = frm1.vspdData.Text
	End If
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""							' Where Condition%>
	arrParam(5) = "근태코드"			' 조건필드의 라벨 명칭 
	
    arrField(0) = "dilig_cd"					' Field명(0)
	arrField(1) = "dilig_nm"					' Field명(1)
	arrField(2) = "day_time"					' Field명(2)
	    
    arrHeader(0) = "근태코드"			' Header명(0)
    arrHeader(1) = "근태코드명"		' Header명(1)
    arrHeader(2) = "일수/시간"		' Header명(2)
 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtCd.focus
		Else
			frm1.vspdData.Col = C_CD
			frm1.vspdData.action = 0
		End If
		Exit Function
	Else
		Call SetAttend(arrRet, iWhere)
	End If	
	
End Function

'------------------------------------------  SetAttend()  --------------------------------------------
'	Name : SetAttend()
'	Description : 근태코드 Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetAttend(Byval arrRet, Byval iWhere)
	Dim strFlag
	Dim lRow
	
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtCd.value = arrRet(0)
			.txtNm.value = arrRet(1)
			.txtCd.focus
		Else 'spread
			.vspdData.Col = C_CD
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_NM
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_FLAG
			.vspdData.Text = arrRet(2)
			
			strFlag = frm1.vspdData.value
			lRow = .vspdData.Row
        
			If strFlag = "1" Then
				.vspdData.Col = C_HOUR
				.vspdData.value = "0"
			    .vspdData.Col = C_MIN
				.vspdData.value = "0"
				ggoSpread.SSSetProtected		C_HOUR, lRow, lRow
				ggoSpread.SSSetProtected		C_MIN, lRow, lRow
			ElseIf strFlag = "2" Then
				ggoSpread.SpreadUnLock          C_HOUR, lRow, C_MIN, lRow
				ggoSpread.SSSetRequired			C_HOUR, lRow, lRow
				ggoSpread.SSSetRequired			C_MIN, lRow, lRow
				
			Else
				.vspdData.Col = C_HOUR
				.vspdData.value = "4"
			    .vspdData.Col = C_MIN
				.vspdData.value = "0"
				ggoSpread.SSSetProtected		C_HOUR, lRow, lRow
				ggoSpread.SSSetProtected		C_MIN, lRow, lRow	
			End If
			.vspdData.Col = C_CD
			.vspdData.action = 0
		End If
	End With
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	ggoSpread.Source = frm1.vspdData
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    
    If Row > 0 Then
		Select Case Col
			Case C_EmpPopup
				Call OpenEmp(1)
			Case C_CdPopup
				Call OpenAttend(1)        
		End Select    
	End If
            
End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )

    Dim iDx
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strValidDt

    Dim strAllColVal
	Dim arrRet
	Dim strFg
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_EMPNO
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_EMPNO
    
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_EMPNM
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DEPT_NM
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DEPT
                Frm1.vspdData.value = ""                
  	            Frm1.vspdData.Col = C_GRD
                Frm1.vspdData.value = ""
            Else
	            IntRetCd = FuncGetEmpInf2(iDx,lgUsrIntCd,strName,strDept_nm, strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	            if  IntRetCd < 0 then
	                if  IntRetCd = -1 then
                		Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                    else
                        Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                    end if
  	                Frm1.vspdData.Col = C_EMPNM
                    Frm1.vspdData.value = ""
  					Frm1.vspdData.Col = C_DEPT_NM
					Frm1.vspdData.value = ""                    
  	                Frm1.vspdData.Col = C_DEPT
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_GRD
                    Frm1.vspdData.value = ""
					vspdData_Change = true
                Else
		       	    Frm1.vspdData.Col = C_EMPNM
		       	    Frm1.vspdData.value = strName
		       	    Frm1.vspdData.Col = C_GRD
		       	    Frm1.vspdData.value = strRoll_pstn

					' 부서정보 가져오는 로직 정리 2003.9.17 by lsn 	
					call SetSpreadDept(iDx)
                End if 
            End if 
         Case  C_CD
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_CD
    
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_CD
                Frm1.vspdData.value = ""
            Else
                IntRetCd = CommonQueryRs(" dilig_cd,DILIG_NM,day_time "," HCA010T "," DILIG_CD =  " & FilterVar(iDx , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
                If IntRetCd = false then
			        Call DisplayMsgBox("800099","X","X","X")	'해당근태는 근태코드정보에 존재하지 않습니다.
  	                Frm1.vspdData.Col = C_NM
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_HOUR 
                    Frm1.vspdData.value = "0"
  	                Frm1.vspdData.Col = C_MIN 
                    Frm1.vspdData.value = "0"
                Else
                    strAllColVal = Trim(Replace(lgF0,Chr(11),"")) & chr(11) & Trim(Replace(lgF1,Chr(11),"")) & chr(11) & Trim(Replace(lgF2,Chr(11),""))
                    arrRet = split(strAllColVal,chr(11))
                    Call SetAttend(arrRet, 1)
                End if 
            End if 
    End Select    
             
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

'========================================================================================================
'   Event Name : txtCd_OnChange
'   Event Desc :
'========================================================================================================
Function txtCd_OnChange()    
    Dim IntRetCd

    If frm1.txtCd.value = "" Then
        frm1.txtNm.value = ""
    ELSE    
        IntRetCd = CommonQueryRs(" dilig_nm "," hca010t "," dilig_cd =  " & FilterVar(frm1.txtCd.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
            Call DisplayMsgBox("800099","X","X","X")                         '☜ : 해당사원은 존재하지 않습니다.            
            frm1.txtNm.value=""
            frm1.txtCd.focus
			txtCd_OnChange = true
			Exit Function
        Else
            frm1.txtNm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
   
End Function

'========================================================================================================
'   Event Name : txtFrDept_Onchange
'========================================================================================================
Function txtFrDept_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
	
    If frm1.txtFrDept.value = "" Then
		frm1.txtFrDeptNm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFrDept.value , frm1.txtValidDt.text , lgUsrIntCd,Dept_Nm , Internal_cd)
    
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFrDeptNm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFrDept.focus
            Set gActiveElement = document.ActiveElement 
            txtFrDept_Onchange = true
            Exit Function      
        Else
			frm1.txtFrDeptNm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function


'========================================================================================================
'   Event Name : txtToDept_Onchange
'========================================================================================================
Function txtToDept_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
	    
    If frm1.txtToDept.value = "" Then
		frm1.txtToDeptNm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtToDept.value , frm1.txtValidDt.text , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtToDeptNm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtToDept.focus
            Set gActiveElement = document.ActiveElement 
            txtToDept_Onchange = true
            Exit Function      
        Else
			frm1.txtToDeptNm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtEmpNo_OnChange
'   Event Desc :
'========================================================================================================
Function txtEmpNo_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmpNo.value = "" Then
		frm1.txtEmpNm.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmpNo.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtEmpNm.value = ""
            Frm1.txtEmpNo.focus 
            Set gActiveElement = document.ActiveElement
			txtEmpNo_Onchange = true
        Else
			frm1.txtEmpNm.value = strName
        End if 
    End if  
End Function
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If  
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If      	
    frm1.vspdData.Row = Row   	
   	
   	
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
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
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

'=======================================================================================================
'   Event Name : txtValidDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtValidDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtValidDt.Action = 7
        frm1.txtValidDt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidDt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtValidDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- space Area-->

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>일일근태등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtValidDt name=txtValidDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="일자"></OBJECT>');</SCRIPT>
								<TD CLASS="TD5" NOWRAP>부서</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtFrDept" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                    <INPUT NAME="txtFrDeptNm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU">&nbsp;~
		                                        <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>근태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCd" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="근태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenAttend(0)">
									            <INPUT TYPE=TEXT NAME="txtNm" tag="14XXXU"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtToDept" MAXLENGTH="10" SIZE=10 ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                    <INPUT NAME="txtToDeptNm" MAXLENGTH="40" SIZE=20 ALT ="Order ID" tag="14XXXU">
    			                                <INPUT NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>사원</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtEmpNo" SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="사원"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCountryCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp(0)">
									            <INPUT TYPE=TEXT NAME="txtEmpNm" tag="14XXXU">
								</TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100% WIDTH=100% >
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

