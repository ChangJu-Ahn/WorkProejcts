<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 승급/승격조회 
*  3. Program ID           : H3007ma1
*  4. Program Name         : H3007ma1
*  5. Program Desc         : 근무이력관리/승급/승격조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/24
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
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
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H3007mb1.asp"                                      'Biz Logic ASP 
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

Dim C_NAME														<%'Spread Sheet의 Column별 상수 %>
Dim C_EMP_NO
Dim C_DEPT_CD 
Dim C_PAY_GRD1 
Dim C_PAY_GRD2 
Dim C_ROLL_PSTN 
Dim C_OCPT_TYPE 
Dim C_FUNC_CD 
Dim C_ROLE_CD
Dim C_ENTR_DT
Dim C_RESENT_PROMOTE_DT
Dim C_CHNG_DEPT_CD
Dim C_CHNG_PAY_GRD1 
Dim C_CHNG_PAY_GRD2 
Dim C_CHNG_ROLL_PSTN
Dim C_CHNG_OCPT_TYPE 
Dim C_CHNG_FUNC_CD 
Dim C_CHNG_ROLE_CD
Dim C_PROMOTE_DT 
Dim C_CHNG_CODE
Dim C_CHNG_CD

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
sub InitSpreadPosVariables()
	C_NAME = 1															<%'Spread Sheet의 Column별 상수 %>
	C_EMP_NO = 2
	C_DEPT_CD = 3
	C_PAY_GRD1 = 4
	C_PAY_GRD2 = 5
	C_ROLL_PSTN = 6
	C_OCPT_TYPE = 7
	C_FUNC_CD = 8
	C_ROLE_CD = 9
	C_ENTR_DT = 10
	C_RESENT_PROMOTE_DT = 11
	C_CHNG_DEPT_CD = 12
	C_CHNG_PAY_GRD1 = 13
	C_CHNG_PAY_GRD2 = 14
	C_CHNG_ROLL_PSTN = 15
	C_CHNG_OCPT_TYPE = 16
	C_CHNG_FUNC_CD = 17
	C_CHNG_ROLE_CD = 18
	C_PROMOTE_DT = 19
	C_CHNG_CODE = 20
	C_CHNG_CD = 21
end sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
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
	frm1.txtPromote_dt1.text =  UniConvDateAToB("<%=GetSvrDate%>",  parent.gServerDateFormat,  parent.gDateFormat)
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
   
    lgKeyStream = Frm1.txtChng_pay_grd11.Value & parent.gColSep                                           'You Must append one character( parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtChng_pay_grd2.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtChng_roll_pstn1.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtChng_cd1.Value & parent.gColSep        ' 변동사유 
    lgKeyStream = lgKeyStream & Frm1.txtOld_dept_cd1.Value & parent.gColSep    ' 발령전부서 
    lgKeyStream = lgKeyStream & Frm1.txtDept_cd1.Value & parent.gColSep        ' 발령후부서 
    lgKeyStream = lgKeyStream & Frm1.txtPromote_dt1.Text & parent.gColSep
    lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep                    ' 자료권한 
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
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
	       .MaxCols = C_CHNG_CD + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:

       .MaxRows = 0
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData        
		Call GetSpreadColumnPos("A")

         ggoSpread.SSSetEdit   C_NAME,             "성명",      10
         ggoSpread.SSSetEdit   C_EMP_NO,           "사번",      13
         ggoSpread.SSSetEdit   C_DEPT_CD,          "부서",      20
         ggoSpread.SSSetEdit   C_PAY_GRD1,         "급호",      10
         ggoSpread.SSSetEdit   C_PAY_GRD2,         "호봉",      10
         ggoSpread.SSSetEdit   C_ROLL_PSTN,        "직위",      20
         ggoSpread.SSSetEdit   C_OCPT_TYPE,        "직종",      20
         ggoSpread.SSSetEdit   C_FUNC_CD,          "직무",      20
         ggoSpread.SSSetEdit   C_ROLE_CD,          "직책",      20
         ggoSpread.SSSetDate   C_ENTR_DT,          "입사일",    10,2,  parent.gDateFormat
         ggoSpread.SSSetDate   C_RESENT_PROMOTE_DT,"최근승급일",10,2,  parent.gDateFormat
         ggoSpread.SSSetEdit   C_CHNG_DEPT_CD,     "변동부서",  15
         ggoSpread.SSSetEdit   C_CHNG_PAY_GRD1,    "변동급호",  10
         ggoSpread.SSSetEdit   C_CHNG_PAY_GRD2,    "변동호봉",  10
         ggoSpread.SSSetEdit   C_CHNG_ROLL_PSTN,   "변동직위",  10
         ggoSpread.SSSetEdit   C_CHNG_OCPT_TYPE,   "변동직종",  10
         ggoSpread.SSSetEdit   C_CHNG_FUNC_CD,     "변동직무",  10
         ggoSpread.SSSetEdit   C_CHNG_ROLE_CD,     "변동직책",  10
         ggoSpread.SSSetDate   C_PROMOTE_DT,       "승급일",    10,2,  parent.gDateFormat
         ggoSpread.SSSetEdit   C_CHNG_CODE,        "변동사유",  10
         ggoSpread.SSSetEdit   C_CHNG_CD,          "변동사유",  10

        Call ggoSpread.SSSetColHidden(C_DEPT_CD,C_DEPT_CD,True)	
        Call ggoSpread.SSSetColHidden(C_CHNG_CODE,C_CHNG_CODE,True)	
        Call ggoSpread.SSSetColHidden(C_PAY_GRD1,C_PAY_GRD1,True)	        
        Call ggoSpread.SSSetColHidden(C_PAY_GRD2,C_PAY_GRD2,True)	
        Call ggoSpread.SSSetColHidden(C_ROLL_PSTN,C_ROLL_PSTN,True)	
        Call ggoSpread.SSSetColHidden(C_OCPT_TYPE,C_OCPT_TYPE,True)	        
        Call ggoSpread.SSSetColHidden(C_FUNC_CD,C_FUNC_CD,True)	        
        Call ggoSpread.SSSetColHidden(C_ROLE_CD,C_ROLE_CD,True)	
        Call ggoSpread.SSSetColHidden(C_ENTR_DT,C_ENTR_DT,True)	
        Call ggoSpread.SSSetColHidden(C_RESENT_PROMOTE_DT,C_RESENT_PROMOTE_DT,True)	        
 
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
            
			C_NAME = iCurColumnPos(1)													<%'Spread Sheet의 Column별 상수 %>
			C_EMP_NO = iCurColumnPos(2)
			C_DEPT_CD = iCurColumnPos(3)
			C_PAY_GRD1 = iCurColumnPos(4)
			C_PAY_GRD2 = iCurColumnPos(5)
			C_ROLL_PSTN = iCurColumnPos(6)
			C_OCPT_TYPE = iCurColumnPos(7)
			C_FUNC_CD = iCurColumnPos(8)
			C_ROLE_CD = iCurColumnPos(9)
			C_ENTR_DT = iCurColumnPos(10)
			C_RESENT_PROMOTE_DT = iCurColumnPos(11)
			C_CHNG_DEPT_CD = iCurColumnPos(12)
			C_CHNG_PAY_GRD1 = iCurColumnPos(13)
			C_CHNG_PAY_GRD2 = iCurColumnPos(14)
			C_CHNG_ROLL_PSTN = iCurColumnPos(15)
			C_CHNG_OCPT_TYPE = iCurColumnPos(16)
			C_CHNG_FUNC_CD = iCurColumnPos(17)
			C_CHNG_ROLE_CD = iCurColumnPos(18)
			C_PROMOTE_DT = iCurColumnPos(19)
			C_CHNG_CODE = iCurColumnPos(20)
			C_CHNG_CD = iCurColumnPos(21)
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
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
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
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
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
    Call SetToolbar("1100101100001111")										        '버튼 툴바 제어 
    
    frm1.txtChng_Pay_grd11.Focus

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
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If

    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    Call  ggoOper.ClearField(Document, "2")                                       '☜: Clear Contents  Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    Call SetToolbar("1100101100001111")
    Call SetDefaultVal
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
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
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       Call UpdateSpreadSheet(Frm1.VspdData.ActiveRow)
    End If    
    
    
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If

    Call  DisableToolBar( parent.TBC_SAVE)
	If DBSave=False Then
	   Call  RestoreToolBar()
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
			SetSpreadColor .ActiveRow, .ActiveRow
    
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
    Set gActiveElement = document.ActiveElement   
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
    	lDelRows =  ggoSpread.DeleteRow
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
    Call parent.FncExport( parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

	If LayerShowHide(1)=False Then
		Exit Function
	End If
	
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
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
        
               Case  ggoSpread.InsertFlag
               Case  ggoSpread.UpdateFlag
               Case  ggoSpread.DeleteFlag                                   

                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	  : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROMOTE_DT  : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_CHNG_CODE   : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1                
           End Select
       Next

       .txtMode.value        =  parent.UID_M0002
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
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If    
    
    Call  DisableToolBar( parent.TBC_DELETE)
	If DbDelete = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

    FncDelete = True                                                        '⊙: Processing is OK


End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     

    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100101100001111")									

    call vspdData_ScriptLeaveCell(1,0,1,1,"")
    frm1.vspddata.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
	
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function


Sub UpdateSpreadSheet(pRow)

     ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow pRow
    With Frm1
           .vspdData.Col  = C_MINORLEN
           .vspdData.Text = .fpdsCdLen.Value
           .vspdData.Col  = C_TYPECd
           .vspdData.Text = .cboType.Value
    End With       

End Sub

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 1 Then
		arrParam(0) = frm1.txtold_Dept_cd1.value
	Else
		arrParam(0) = frm1.txtDept_cd1.value
	End If
	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 1 Then
			frm1.txtold_Dept_cd1.focus
		Else
			frm1.txtDept_cd1.focus
		End If	
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
		If iWhere = 1 Then
			.txtold_Dept_cd1.value = arrRet(0)
			.txtold_Dept_cd_nm.value = arrRet(1)
			.txtold_Dept_cd1.focus
		Else
			.txtDept_cd1.value = arrRet(0)
			.txtDept_cd_nm.value = arrRet(1)
			.txtDept_cd1.focus
		End If
	End With
End Function

'===========================================================================
' Function Name : OpenSItemDC
' Function Desc : OpenSItemDC Reference Popup
'===========================================================================
Function OpenSItemDC(iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1  ' 급호 
	    	arrParam(0) = "급호코드 팝업"		            ' 팝업 명칭 
	    	arrParam(1) = "B_minor"				              	' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtChng_Pay_grd11.Value)	        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""		    		' Where Condition
	    	arrParam(5) = "급호코드"	   				    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)%>
    
	    	arrHeader(0) = "급호코드"		        		' Header명(0)%>
	    	arrHeader(1) = "급호명"	       					' Header명(1)%>

	    Case 2  ' 직위 
	    	arrParam(0) = "직위코드 팝업"		            ' 팝업 명칭 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtChng_Roll_pstn1.Value)	' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0002", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직위코드"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직위코드"			        	' Header명(0)
	    	arrHeader(1) = "직위명"	        				' Header명(1)

	    Case 29  ' 사유 
	        arrParam(0) = "사유코드 팝업"		        ' 팝업 명칭 
	        arrParam(1) = "B_minor"				 	        ' TABLE 명칭 
	        arrParam(2) = Trim(frm1.txtchng_cd1.Value)	    ' Code Condition
	        arrParam(3) = ""						        ' Name Cindition
	        arrParam(4) = "major_cd=" & FilterVar("H0029", "''", "S") & ""                ' Where Condition
	        arrParam(5) = "사유코드"			
	
            arrField(0) = "minor_cd"				        ' Field명(0)
            arrField(1) = "minor_nm"				        ' Field명(1)
    
            arrHeader(0) = "사유코드"			        ' Header명(0)
            arrHeader(1) = "사유명"			            ' Header명(1)
    
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	
		Select Case iWhere
		    Case 1
		    	frm1.txtChng_pay_grd11.focus
		    Case 2
		    	frm1.txtChng_Roll_pstn1.focus
		    Case 29
		    	frm1.txtchng_cd1.focus
		End Select
		Exit Function
	Else
		Call SetSItemDC(arrRet, iWhere)
	End If	
	
End Function
'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetSItemDC()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSItemDC(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case 1
		    	.txtChng_pay_grd11.value = arrRet(0)
		    	.txtChng_Pay_grd11_nm.value = arrRet(1)  
		    	.txtChng_pay_grd11.focus
		    Case 2
		    	.txtChng_Roll_pstn1.value = arrRet(0)
		    	.txtChng_Roll_pstn1_nm.value = arrRet(1)
		    	.txtChng_Roll_pstn1.focus
		    Case 29
		    	.txtchng_cd1.value = arrRet(0)
		    	.txtchng_cd1_nm.value = arrRet(1)
		    	.txtchng_cd1.focus
		End Select

		lgBlnFlgChgValue = True

	End With
	
End Function

'========================================================================================================
' Name : OpenPromoteDt
' Desc : 최근 승급일 POPUP
'========================================================================================================
Function OpenPromoteDt(iWhere)
	Dim arrRet
	Dim arrParam(1)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtPromote_dt1.text			
	arrParam(1) = 2 'hba080t
	arrRet = window.showModalDialog(HRAskPRAspName("PromoteDtPopup"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	if arrRet(0) <> ""	then
		frm1.txtPromote_dt1.text = arrRet(0)
	end if
		frm1.txtPromote_dt1.focus	
End Function




Sub txtChng_pay_grd11_OnChange()

    If  frm1.txtChng_pay_grd11.value <> "" Then
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0001", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtChng_pay_grd11.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtChng_pay_grd11_nm.value = ""
            Call  DisplayMsgBox("970000", "x","급호코드","x")
	        frm1.txtChng_pay_grd11.focus
	        Set gActiveElement = document.ActiveElement
	    Else
	        frm1.txtChng_pay_grd11_nm.value = Replace(lgF0, Chr(11), "")
	    End If
	Else
            frm1.txtChng_pay_grd11_nm.value = ""			    
    End If
	    
End Sub

Sub txtChng_Roll_pstn1_OnChange()

    If  frm1.txtChng_Roll_pstn1.value <> "" Then
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0002", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtChng_Roll_pstn1.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtChng_roll_pstn1_nm.value = ""
            Call  DisplayMsgBox("970000", "x","직위코드","x")
	        frm1.txtChng_roll_pstn1.focus
	        Set gActiveElement = document.ActiveElement
	    Else
	        frm1.txtChng_roll_pstn1_nm.value = Replace(lgF0, Chr(11), "")
	    End If
	Else
            frm1.txtChng_roll_pstn1_nm.value = ""		    
    End If
	    
End Sub

Sub txtChng_cd1_OnChange()

    If  frm1.txtChng_cd1.value <> "" Then
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0029", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtChng_cd1.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtChng_cd1_nm.value = ""
            Call  DisplayMsgBox("970000", "x","사유코드","x")
	        frm1.txtChng_cd1.focus
	        Set gActiveElement = document.ActiveElement
	    Else
	        frm1.txtChng_cd1_nm.value = Replace(lgF0, Chr(11), "")
	    End If
	Else	   
            frm1.txtChng_cd1_nm.value = ""	 
    End If
	    
End Sub

Sub txtold_Dept_cd1_OnChange()

    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd

    if  frm1.txtold_Dept_cd1.value = "" then
        frm1.txtold_Dept_cd_nm.value = ""
    else
        IntRetCd =  FuncDeptName(frm1.txtold_Dept_cd1.value,frm1.txtPromote_dt1.text,lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            frm1.txtold_Dept_cd_nm.value = ""            
            frm1.txtold_Dept_cd1.focus
	        Set gActiveElement = document.ActiveElement
        else
            frm1.txtold_Dept_cd_nm.value = strDept_nm
        end if
    end if

End Sub

Sub txtDept_cd1_OnChange()

    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd
    
    if  frm1.txtDept_cd1.value = "" then
        frm1.txtDept_cd_nm.value = ""
    else
        IntRetCd =  FuncDeptName(frm1.txtDept_cd1.value,frm1.txtPromote_dt1.text,lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            frm1.txtDept_cd_nm.value = ""              
            frm1.txtDept_cd1.focus
	        Set gActiveElement = document.ActiveElement
        else
            frm1.txtDept_cd_nm.value = strDept_nm
        end if
    end if

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
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

	Call SetPopupMenuItemInf("0101111111")       

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
	call vspdData_ScriptLeaveCell(Col,Row-1,Col,Row,"")
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
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
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
 
    Dim iRet

	If Not (Row <> NewRow And NewRow > 0) Then    
	   Exit Sub
	End If

    If lgBlnFlgChgValue = True Then 
       Call UpdateSpreadSheet(Row)
       lgBlnFlgChgValue = False
    End If

	With Frm1
        .vspdData.Row = NewRow

        .vspdData.Col = C_DEPT_CD
        .txtDept_cd.Value = .vspdData.Text
        .vspdData.Col = C_PAY_GRD1
        .txtPay_grd1.Value = .vspdData.Text
        .vspdData.Col = C_PAY_GRD2
        .txtPay_grd2.Value = .vspdData.Text
        .vspdData.Col = C_ROLL_PSTN
        .txtRoll_pstn.Value = .vspdData.Text
        .vspdData.Col = C_OCPT_TYPE
        .txtOcpt_type.Value = .vspdData.Text
        .vspdData.Col = C_FUNC_CD
        .txtFunc_cd.Value = .vspdData.Text
        .vspdData.Col = C_ROLE_CD
        .txtRole_cd.Value = .vspdData.Text

        .vspdData.Col = C_CHNG_DEPT_CD
        .txtChng_Dept_cd.Value = .vspdData.Text
        .vspdData.Col = C_CHNG_PAY_GRD1
        .txtChng_Pay_grd1.Value = .vspdData.Text
        .vspdData.Col = C_CHNG_PAY_GRD2
        .txtChng_Pay_grd2.Value = .vspdData.Text
        .vspdData.Col = C_CHNG_ROLL_PSTN
        .txtCHNG_Roll_pstn.Value = .vspdData.Text
        .vspdData.Col = C_CHNG_OCPT_TYPE
        .txtCHNG_Ocpt_type.Value = .vspdData.Text
        .vspdData.Col = C_CHNG_FUNC_CD
        .txtCHNG_Func_cd.Value = .vspdData.Text
        .vspdData.Col = C_CHNG_ROLE_CD
        .txtCHNG_Role_cd.Value = .vspdData.Text

        .vspdData.Col = C_ENTR_DT
        .txtEntr_dt.Value =  .vspdData.Text
        .vspdData.Col = C_CHNG_CD
        .txtChng_cd.Value =  .vspdData.Text
        .vspdData.Col = C_PROMOTE_DT
        .txtPromote_dt.Value =  .vspdData.Text
    End With       
           
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


Sub cboType_OnChange()
    lgBlnFlgChgValue = True
End Sub

Sub fpdsCdLen_KeyDown(KeyCode , Shift)
    lgBlnFlgChgValue = True
End Sub 

'========================================================================================================
' Name : txtPromote_dt1_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtPromote_dt1_DblClick(Button)
    If Button = 1 Then
   		Call SetFocusToDocument("M") 
        frm1.txtPromote_dt1.Action = 7 
        frm1.txtPromote_dt1.focus
    End If
    lgBlnFlgChgValue = True    
End Sub

Sub txtPromote_dt1_Change()
    lgBlnFlgChgValue = True    
End Sub

'==========================================================================================
'   Event Name : txtPromote_dt1_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtPromote_dt1_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>승급/승격조회</font></td>
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
								<TD CLASS=TD5 NOWRAP>변동급호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChng_Pay_grd11" ALT="변동급호" TYPE="Text" MAXLENGTH=10 SiZE=10 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC(1)">&nbsp;<INPUT NAME="txtChng_Pay_grd11_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
								</TD>
								<TD CLASS=TD5 NOWRAP>변동직위</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChng_Roll_pstn1" ALT="변동직위" TYPE="Text" MAXLENGTH=10 SiZE=10 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC(2)">&nbsp;<INPUT NAME="txtChng_Roll_pstn1_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>승급일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h3007ma1_txtPromote_dt1_txtPromote_dt1.js'></script>
								<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPromoteDt" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPromoteDt(0)"></TD>
								<TD CLASS=TD5 NOWRAP>변동사유</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtChng_cd1" ALT="변동사유" TYPE="Text" MAXLENGTH=10 SiZE=10 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC(29)">&nbsp;<INPUT NAME="txtChng_cd1_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발령전부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtold_Dept_cd1" ALT="부서" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept(1)">&nbsp;<INPUT NAME="txtold_Dept_cd_nm" TYPE="Text" MAXLENGTH="30" SIZE=20 tag="14">
								</TD>
								<TD CLASS=TD5 NOWRAP>발령후부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd1" ALT="부서" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept(2)">&nbsp;<INPUT NAME="txtDept_cd_nm" TYPE="Text" MAXLENGTH="30" SIZE=20 tag="14">
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN=4>
									<script language =javascript src='./js/h3007ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
			            </TABLE>
                    </TD>
				</TR>
				<TR>
                	<TD HEIGHT=* WIDTH=100%>
                		<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
	            				<TR>
	            					<TD CLASS="TD5" NOWRAP></TD>
	            					<TD CLASS="TD6" NOWRAP>발령이전&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	            					                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;발령이후</TD>
	            					<TD CLASS="TD5" NOWRAP></TD>
	            					<TD CLASS="TD6" NOWRAP></TD>                    					
	            				</TR>
	            				<TR>
	            					<TD CLASS="TD5" NOWRAP>부서</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtDept_cd" Size="20" MAXLENGTH="10" ALT="부서" Tag="24">
	            					<INPUT TYPE=TEXT Name="txtChng_dept_cd" Size="20" MAXLENGTH="10" ALT="부서2" Tag="24">
	            					</TD>
	            					<TD CLASS="TD5" NOWRAP>변동사유</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtChng_cd" Size="15" MAXLENGTH="10" ALT="변동사유" Tag="24">
	            					</TD>
	            				</TR>
	            				<TR>
	            					<TD CLASS="TD5" NOWRAP>급호</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtPay_grd1" Size="13" MAXLENGTH="10" ALT="급호" Tag="24">&nbsp;<INPUT NAME="txtPay_grd2" MAXLENGTH="5" SIZE=5 STYLE="TEXT-ALIGN:left" tag="24">
	            					<INPUT TYPE=TEXT Name="txtChng_pay_grd1" Size="13" MAXLENGTH="10" ALT="급호2" Tag="24">&nbsp;<INPUT NAME="txtChng_pay_grd2" MAXLENGTH="5" SIZE=5 STYLE="TEXT-ALIGN:left" tag="24">
	            					</TD>
	            					<TD CLASS="TD5" NOWRAP>입사일</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtEntr_dt" Size="15" MAXLENGTH="10" ALT="입사일" Tag="24">
	            					</TD>
	            				</TR>
	            				<TR>
	            					<TD CLASS="TD5" NOWRAP>직위</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtRoll_pstn" Size="20" MAXLENGTH="10" ALT="직위" Tag="24">
	            					<INPUT TYPE=TEXT Name="txtChng_roll_pstn" Size="20" MAXLENGTH="10" ALT="직위2" Tag="24">
	            					</TD>
	            					<TD CLASS="TD5" NOWRAP>승급일</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtPromote_dt" Size="15" MAXLENGTH="10" ALT="승급일" Tag="24">
	            					</TD>
	            				</TR>
	            				<TR>
	            					<TD CLASS="TD5" NOWRAP>직종</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtOcpt_type" Size="20" MAXLENGTH="10" ALT="직종" Tag="24">
	            					<INPUT TYPE=TEXT Name="txtChng_ocpt_type" Size="20" MAXLENGTH="10" ALT="직종2" Tag="24">
	            					</TD>
	            					<TD CLASS="TD5" NOWRAP></TD>
	            					<TD CLASS="TD6" NOWRAP></TD>
	            				</TR>
	            				<TR>
	            					<TD CLASS="TD5" NOWRAP>직무</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtFunc_cd" Size="20" MAXLENGTH="10" ALT="직무" Tag="24">
	            					<INPUT TYPE=TEXT Name="txtChng_func_cd" Size="20" MAXLENGTH="10" ALT="직무2" Tag="24">
	            					</TD>
	            					<TD CLASS="TD5" NOWRAP></TD>
	            					<TD CLASS="TD6" NOWRAP></TD>
	            				</TR>
	            				<TR>
	            					<TD CLASS="TD5" NOWRAP>직책</TD>
	            					<TD CLASS="TD6" NOWRAP>
	            					<INPUT TYPE=TEXT Name="txtRole_cd" Size="20" MAXLENGTH="10" ALT="직책" Tag="24">
	            					<INPUT TYPE=TEXT Name="txtChng_role_cd" Size="20" MAXLENGTH="10" ALT="직책2" Tag="24">
	            					</TD>
	            					<TD CLASS="TD5" NOWRAP></TD>
	            					<TD CLASS="TD6" NOWRAP></TD>
	            				</TR>
				        </TABLE>
				    </TD>
				</TR>
            </TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

