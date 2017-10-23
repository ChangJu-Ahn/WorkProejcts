<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Multi Sample
*  3. Program ID           : H5102ma1
*  4. Program Name         : H5102ma1
*  5. Program Desc         : 고정공제사항등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/01/02
*  8. Modified date(Last)  : 2002/01/02
*  9. Modifier (First)     : chcho
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h5102mb1.asp"	
Const BIZ_PGM_ID1     = "h5102mb2.asp"						           '☆: Biz Logic ASP Name
Const CookieSplit = 1233
Const C_SHEETMAXROWS    =   21	                                      '한 화면에 보여지는 최대갯수*1.5%>
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop 
Dim lgStrComDateType		'Company Date Type을 저장(년월 Mask에 사용함.)

Dim C_EMP_NO
Dim C_NAME_POP
Dim C_NAME
Dim C_EMP_NO_POP
Dim C_DEPT_CD   
Dim C_DEPT_NM   
Dim C_SUB_TYPE  
Dim C_SUB_TYPE_NM
Dim C_SUB_TYPE_POP
Dim C_SUB_CD      
Dim C_SUB_CD_NM   
Dim C_SUB_CD_POP  
Dim C_SUB_AMT     
Dim C_APPLY_YYMM  
Dim C_REVOKE_YYMM 

'========================================================================================================
' Name : InitSpreadPosVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_EMP_NO         = 1	
	 C_NAME_POP       = 2
	 C_NAME           = 3
	 C_EMP_NO_POP     = 4
	 C_DEPT_CD        = 5
	 C_DEPT_NM        = 6
	 C_SUB_TYPE       = 7
	 C_SUB_TYPE_NM    = 8
	 C_SUB_TYPE_POP   = 9
	 C_SUB_CD         = 10
	 C_SUB_CD_NM      = 11 
	 C_SUB_CD_POP     = 12 
	 C_SUB_AMT        = 13
	 C_APPLY_YYMM     = 14
	 C_REVOKE_YYMM    = 15
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
	
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)

	frm1.txtapply_yymm_dt.Year = strYear 
	frm1.txtapply_yymm_dt.Month = strMonth 

	frm1.txtrevoke_yymm_dt.Year = strYear 
	frm1.txtrevoke_yymm_dt.Month = strMonth 
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
Sub MakeKeyStream(pOpt)
   
       lgKeyStream  = Trim(Frm1.txtemp_No.Value) & parent.gColSep                        '0
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtsub_type.Value) & parent.gColSep        '1
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtsub_cd.Value) & parent.gColSep          '2
      
       lgKeyStream  = lgKeyStream & frm1.txtFr_internal_cd.value & parent.gColSep  '3
       lgKeyStream  = lgKeyStream & frm1.txtTo_internal_cd.value & parent.gColSep  '4
       
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtapply_yymm_dt.text) & parent.gColSep    '5
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtrevoke_yymm_dt.text) & parent.gColSep   '6
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtsub_type_nm.Value) & parent.gColSep     '7
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtsub_cd_nm.Value) & parent.gColSep       '8
 
       lgKeyStream  = lgKeyStream & lgUsrIntcd & parent.gColSep  '9 
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

    Dim strMaskYM	

	If Date_DefMask(strMaskYM) = False Then
		strMaskYM = "9999" & lgStrComDateType & "99"
	End If	

	Call initSpreadPosVariables()	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_REVOKE_YYMM + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0	
        ggoSpread.ClearSpreadData
       
		 Call  GetSpreadColumnPos("A")
       
         ggoSpread.SSSetEdit     C_EMP_NO        , "사번", 13,,, 13,2
         ggoSpread.SSSetButton   C_NAME_POP
         ggoSpread.SSSetEdit     C_NAME          , "성명", 12,,, 30,2         
         ggoSpread.SSSetButton   C_EMP_NO_POP
         ggoSpread.SSSetEdit     C_DEPT_CD       , "부서명", 12,,, 20,2  '부서HIDDEN
         ggoSpread.SSSetEdit     C_DEPT_NM       , "부서명", 17,,, 40,2
         ggoSpread.SSSetEdit     C_SUB_TYPE      , "공제구분", 12 ,,,1,2'구분HIDDEN
         ggoSpread.SSSetEdit     C_SUB_TYPE_NM   , "공제구분", 11 ,,,50,2
         ggoSpread.SSSetButton   C_SUB_TYPE_POP           
         ggoSpread.SSSetEdit     C_SUB_CD        , "공제코드", 14 ,,,3,2 '코드HIDDEN
         ggoSpread.SSSetEdit     C_SUB_CD_NM     , "공제코드", 15 ,,,50,2
         ggoSpread.SSSetButton   C_SUB_CD_POP         
         ggoSpread.SSSetFloat    C_SUB_AMT       , "공제금액", 18,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
         ggoSpread.SSSetMask     C_APPLY_YYMM    , "적용년월", 10,2, strMaskYM
         ggoSpread.SSSetMask     C_REVOKE_YYMM   , "해제년월", 10,2, strMaskYM
				
		 Call ggoSpread.MakePairsColumn(C_EMP_NO,	   C_NAME_POP)
		 Call ggoSpread.MakePairsColumn(C_SUB_TYPE_NM, C_SUB_TYPE_POP)
		 Call ggoSpread.MakePairsColumn(C_SUB_CD_NM,   C_SUB_CD_POP)
		 
         Call ggoSpread.SSSetColHidden(C_EMP_NO_POP	,  C_EMP_NO_POP	, True)
         Call ggoSpread.SSSetColHidden(C_DEPT_CD	,  C_DEPT_CD	, True)
         Call ggoSpread.SSSetColHidden(C_SUB_TYPE	,  C_SUB_TYPE	, True)
         Call ggoSpread.SSSetColHidden(C_SUB_CD		,  C_SUB_CD		, True)
        
        .ReDraw = true
	 
        Call SetSpreadLock 
    
    End With
    
 
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
      ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
   
       ggoSpread.SpreadLock      C_NAME			, -1, C_NAME
       ggoSpread.SpreadLock      C_NAME_POP		, -1, C_NAME_POP
       ggoSpread.SpreadLock      C_EMP_NO		, -1, C_EMP_NO
       ggoSpread.SpreadLock      C_EMP_NO_POP	, -1, C_EMP_NO_POP
       ggoSpread.SpreadLock      C_DEPT_CD		, -1, C_DEPT_CD
       ggoSpread.SpreadLock      C_DEPT_NM		, -1, C_DEPT_NM
       ggoSpread.SpreadLock      C_SUB_TYPE		, -1, C_SUB_TYPE
       ggoSpread.SpreadLock      C_SUB_TYPE_POP , -1, C_SUB_TYPE_POP
       ggoSpread.SpreadLock      C_SUB_TYPE_NM	, -1, C_SUB_TYPE_NM
       ggoSpread.SpreadLock      C_SUB_CD		, -1, C_SUB_CD
       ggoSpread.SpreadLock      C_SUB_CD_POP	, -1, C_SUB_CD_POP
       ggoSpread.SpreadLock      C_SUB_CD_NM	, -1, C_SUB_CD_NM
       ggoSpread.SSSetRequired	 C_SUB_AMT		, -1, -1
       ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
     .vspdData.ReDraw = True
    END With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False    
    
       ggoSpread.Source = frm1.vspdData  
       ggoSpread.SSSetProtected   C_NAME		, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_EMP_NO		, pvStartRow, pvEndRow
       ggoSpread.SSSetProtected   C_DEPT_CD     , pvStartRow, pvEndRow
       ggoSpread.SSSetProtected   C_DEPT_NM     , pvStartRow, pvEndRow
       ggoSpread.SSSetProtected   C_SUB_TYPE	, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_SUB_TYPE_NM , pvStartRow, pvEndRow
       ggoSpread.SSSetProtected   C_SUB_CD		, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_SUB_CD_NM	, pvStartRow, pvEndRow
       ggoSpread.SSSetRequired    C_SUB_AMT     , pvStartRow, pvEndRow       

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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                
            
            C_EMP_NO         = iCurColumnPos(1)	
			C_NAME_POP       = iCurColumnPos(2)
			C_NAME           = iCurColumnPos(3)
			C_EMP_NO_POP     = iCurColumnPos(4)
			C_DEPT_CD        = iCurColumnPos(5)
			C_DEPT_NM        = iCurColumnPos(6)
			C_SUB_TYPE       = iCurColumnPos(7)
			C_SUB_TYPE_NM    = iCurColumnPos(8)
			C_SUB_TYPE_POP   = iCurColumnPos(9)
			C_SUB_CD         = iCurColumnPos(10)
			C_SUB_CD_NM      = iCurColumnPos(11) 
			C_SUB_CD_POP     = iCurColumnPos(12) 
			C_SUB_AMT        = iCurColumnPos(13)
			C_APPLY_YYMM     = iCurColumnPos(14)
			C_REVOKE_YYMM    = iCurColumnPos(15)                        
            
    End Select    
End Sub
'======================================================================================================
' Function Name : vspdData_ScriptLeaveCell
' Function Desc : 년(YYYY).월(MM) check
'======================================================================================================
Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format
		
    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call  InitSpreadSheet                                                            'Setup the Spread sheet
    Call  InitVariables                                                              'Initializes local global variables
    Call  ggoOper.FormatDate(frm1.txtapply_yymm_dt,  parent.gDateFormat, 2) 
    Call  ggoOper.FormatDate(frm1.txtrevoke_yymm_dt,  parent.gDateFormat, 2) 
    Call  FuncGetAuth(gStrRequestMenuID ,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
	Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 

	frm1.txtsub_type.focus 
    Call CookiePage (0)
   
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
    Dim strwhere
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData

    If  txtsub_type_OnChange() then
        Exit Function
    End If
   
    If  txtsub_cd_Onchange() then
        Exit Function
    End If
    If  txtEmp_no_Onchange()  then
        Exit Function
    End If

    If  txtFr_Dept_cd_Onchange() then
        Exit Function
    End If
    If  txtTo_Dept_cd_Onchange() then
        Exit Function
    End If

    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept, strFrYymm, strToYymm
    
    Fr_dept_cd = Trim(frm1.txtFr_internal_cd.value)
    To_dept_cd = Trim(frm1.txtTo_internal_cd.value)
    
    If fr_dept_cd = "" then
        IntRetCd =  FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd =  FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call  DisplayMsgBox("800359","X","X","X")	'시작부서보다 작은값입니다.
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF   

    If (frm1.txtapply_yymm_dt.Text = "") Then                       '년월의 값이 없으면 주는 기본값정의와 메시지 체크 
        strFrYymm =  UniConvYYYYMMDDToDate( parent.gDateFormat, "1900", "01", "01")
    Else
        strFrYymm =  UniConvYYYYMMDDToDate( parent.gDateFormat, frm1.txtapply_yymm_dt.Year, Right("0" & frm1.txtapply_yymm_dt.month , 2), "01")
    End if 
    
    If (frm1.txtrevoke_yymm_dt.Text = "") Then
        strToYymm =  UniConvYYYYMMDDToDate( parent.gDateFormat, "2500", "12", "31")
    Else
        strToYYMM =  UniConvYYYYMMDDToDate( parent.gDateFormat, frm1.txtrevoke_yymm_dt.Year, Right("0" & frm1.txtrevoke_yymm_dt.month , 2), "01")
    End if 
         
    If  CompareDateByFormat(strFrYymm,strToYymm,frm1.txtapply_yymm_dt.Alt,frm1.txtrevoke_yymm_dt.Alt,"970025", parent.gDateFormat, parent.gComDateType,True) = False Then
        frm1.txtapply_yymm_dt.focus
        Set gActiveElement = document.activeElement

        Exit Function
    End if 

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    If DbQuery = False Then
       Exit Function
    End If                                                                 '☜: Query db data
       
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
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    Dim strApplyDt                                                              '저장시 그리드에입력된 년월의 타당성 체크 
   	Dim strRevokeDt
   	Dim lRow
   	Dim vspd_data
   	Dim vspd_data1
   	Dim test,test1

		With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case  ggoSpread.InsertFlag,  ggoSpread.UpdateFlag               
					.vspdData.Col = C_NAME

					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					    Call DisplayMsgBox("800048","X","X","X")
						Exit Function
					end if
					.vspdData.Col = C_SUB_TYPE

					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					    Call DisplayMsgBox("970000","X","공제구분","X")
						Exit Function
					end if		
					.vspdData.Col = C_SUB_CD
					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					    Call DisplayMsgBox("970000","X","공제코드","X")
						Exit Function
					end if									
   	                .vspdData.Col = C_APPLY_YYMM     	                

					if Trim(.vspdData.Text) = parent.gComDateType  then
						.vspdData.Text = ""
						vspd_Data = ""
					else
						vspd_data = .vspdData.Text
					end if							
   	                
                    strApplyDt =  UNIGetLastDay(.vspdData.Text,  parent.gDateFormatYYYYMM)
                     
                    if (Trim(StrApplyDt) = "" AND vspd_data <> "")  then
                    
						Call DisplayMsgBox("200006","X","X","X")	
	                     .vspdData.focus
                         Set gActiveElement = document.activeElement
                         Exit Function
					end if
					
   	                .vspdData.Col = C_REVOKE_YYMM
   	                
   	               	if Trim(.vspdData.Text) = parent.gComDateType then
						.vspdData.Text = ""
						vspd_data1 = ""
					else
						vspd_data1 = .vspdData.Text
					end if	
   	                
                    strRevokeDt =  UNIGetLastDay(.vspdData.Text,  parent.gDateFormatYYYYMM)
                    
                    if (Trim(strRevokeDt) = "" AND vspd_data1 <> "") then
                    
						Call DisplayMsgBox("200006","X","X","X")	
	                     .vspdData.focus
                         Set gActiveElement = document.activeElement
                         Exit Function
					end if
                    If .vspdData.Text = "" Then
                    Else
                        If  CompareDateByFormat(strApplyDt,strRevokeDt,"적용년월","해제년월","970023", parent.gDateFormat, parent.gComDateType,True) = False then
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_APPLY_YYMM
                            .vspdData.focus
                            Set gActiveElement = document.activeElement
                            Exit Function
                        Else
                        End if 
                    End if  
            End Select
        Next
	End With

	Call  DisableToolBar( parent.TBC_SAVE)
	If DbSAVE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
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
 
            .ReDraw = True
            .Col = C_SUB_CD
		    .Focus
		    .Action = 0 ' go to 
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
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1         
        
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    
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
    	lDelRows =  ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( parent.C_MULTI)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_MULTI, False)
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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
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

	if LayerShowHide(1) = false then
	exit Function
	end if

	Dim strVal
	
    With frm1
       ggoSpread.Source = frm1.vspdData
     .vspdData.ReDraw = False
       ggoSpread.SpreadLock      C_NAME_POP , -1, C_NAME_POP
       ggoSpread.SpreadLock      C_EMP_NO_POP , -1, C_EMP_NO_POP
       ggoSpread.SpreadLock      C_SUB_TYPE_POP , -1, C_SUB_TYPE_POP
       ggoSpread.SpreadLock      C_SUB_CD_POP , -1, C_SUB_CD_POP
     .vspdData.ReDraw = True
    END With

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	if LayerShowHide(1) = false then
	exit Function
	end if
		    
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
' Desc : This function is called by FncSave
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

    Dim iColSep, iRowSep
    Dim strCUTotalvalLen					'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen						'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
 	Dim iFormLimitByte						'102399byte
 	Dim objTEXTAREA							'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 
 	Dim iTmpCUBuffer						'현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount					'현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount				'현재의 버퍼 Chunk Size
 	Dim iTmpDBuffer							'현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount					'현재의 버퍼 Position
	Dim iTmpDBufferMaxCount					'현재의 버퍼 Chunk Size

    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
 	
 	'한번에 설정한 버퍼의 크기 설정 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
 	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
     
     '102399byte
     iFormLimitByte = parent.C_FORM_LIMIT_BYTE
     
     '버퍼의 초기화 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
 	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				
 
 	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
 	
 	strCUTotalvalLen = 0 : strDTotalvalLen  = 0
	
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
	exit Function
	end if

    lGrpCnt = 1

	With Frm1

       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
                  Case  ggoSpread.InsertFlag                                      '☜: Update추가 
					strVal = ""                  
                                                    strVal = strVal & "C" & parent.gColSep 'array(0)
                                                    strVal = strVal & lRow & parent.gColSep
                                                
                    .vspdData.Col = C_NAME	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_TYPE      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_CD        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_AMT   	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_APPLY_YYMM    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REVOKE_YYMM   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    
                    
                    lGrpCnt = lGrpCnt + 1 
               
               Case  ggoSpread.UpdateFlag                                      '☜: Update
					strVal = ""               
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                                                
                    
                    .vspdData.Col = C_NAME	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_TYPE      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_CD        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_AMT   	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_APPLY_YYMM    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REVOKE_YYMM   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
              
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
				    strDel = ""
                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                                                                     
                    .vspdData.Col = C_NAME	        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD       : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_TYPE      : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_SUB_CD        : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                   
                    
                    lGrpCnt = lGrpCnt + 1
           End Select
			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
			                            
			            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
			            objTEXTAREA.name = "txtCUSpread"
			            objTEXTAREA.value = Join(iTmpCUBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			 
			            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
			            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
			            iTmpCUBufferCount = -1
			            strCUTotalvalLen  = 0
			         End If

			         iTmpCUBufferCount = iTmpCUBufferCount + 1
			      
			         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
			            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
			            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
			         End If   

			         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)

			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '한개의 form element에 넣을 한개치가 넘으면 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
           
       Next
	
		.txtMode.value        =  parent.UID_M0002
		.txtUpdtUserId.value  =  parent.gUsrID
		.txtInsrtUserId.value =  parent.gUsrID
		.txtMaxRows.value     = lGrpCnt-1	
	End With

    If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

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
	If DbDELETE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
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
	Call SetToolbar("110011110011111")	 
	frm1.vspdData.focus
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    ggoSpread.Source = frm1.vspdData	
	ggoSpread.ClearSpreadData
	Call RemovedivTextArea	
    Call InitVariables															'⊙: Initializes local global variables
	call MainQuery()
	
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()       
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	     
        Case "4"
            arrParam(0) = "공제구분 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_type.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtsub_type_nm.value							' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " "    ' Where Condition							' Where Condition
	        arrParam(5) = "공제구분코드"			    ' TextBox 명칭 
	
            arrField(0) = "MINOR_CD"					' Field명(0)
            arrField(1) = "MINOR_NM"				    ' Field명(1)
    
            arrHeader(0) = "공제구분코드"				' Header명(0)
            arrHeader(1) = "공제구분명"
	    Case "5"
            arrParam(0) = "공제코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_cd.value		    ' Code Condition
	        arrParam(3) = ""'frm1.txtsub_cd_nm.value			' Name Cindition
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("2", "''", "S") & " "  ' Where Condition
	        arrParam(5) = "공제코드"			    ' TextBox 명칭 
	
            arrField(0) = "ALLOW_CD"					' Field명(0)
            arrField(1) = "ALLOW_NM"				    ' Field명(1)
    
            arrHeader(0) = "공제코드"				' Header명(0)
            arrHeader(1) = "공제코드명"
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "4"
		        frm1.txtsub_type.focus
		    Case "5"
		        frm1.txtsub_cd.focus
        End Select	
		Exit Function
	Else

		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()           
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere) 
	With Frm1
		Select Case iWhere

		    Case "4"
		        .txtsub_type.value = arrRet(0)
		        .txtsub_type_NM.value = arrRet(1)
		        .txtsub_type.focus
		    Case "5"
		        .txtsub_cd.value = arrRet(0)
		        .txtsub_cd_NM.value = arrRet(1)
		        .txtsub_cd.focus
        End Select
	End With
End Sub


'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere

	     Case C_SUB_TYPE_POP
            arrParam(0) = "공제구분 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_type.value		    ' Code Condition
	        arrParam(3) = ""
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " "    ' Where Condition							' Where Condition
	        arrParam(5) = "공제구분"			    ' TextBox 명칭 
	
            arrField(0) = "MINOR_CD"					' Field명(0)
            arrField(1) = "MINOR_NM"				    ' Field명(1)
    
            arrHeader(0) = "공제구분코드"				' Header명(0)
            arrHeader(1) = "공제구분명"
	 
	    Case C_SUB_CD_POP

	        arrParam(0) = "공제코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_cd.value		    ' Code Condition
	        arrParam(3) = ""
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("2", "''", "S") & " "  ' Where Condition
	        arrParam(5) = "공제코드"			    ' TextBox 명칭 
	
            arrField(0) = "ALLOW_CD"					' Field명(0)
            arrField(1) = "ALLOW_NM"				    ' Field명(1)
    
            arrHeader(0) = "공제코드"				' Header명(0)
            arrHeader(1) = "공제코드명"
	   
	    
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case C_SUB_TYPE_POP
		    	frm1.vspdData.Col = C_SUB_TYPE_NM
				frm1.vspdData.action =0
		    Case C_SUB_CD_POP
		    	frm1.vspdData.Col = C_SUB_CD_NM
				frm1.vspdData.action =0
        End Select	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow Row
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere

		    Case C_SUB_TYPE_POP
		    	.vspdData.Col = C_SUB_TYPE
		    	.vspdData.text = arrRet(0)  
		    	.vspdData.Col = C_SUB_TYPE_NM
		    	.vspdData.text = arrRet(1)  
				.vspdData.action =0
		    Case C_SUB_CD_POP
		    	.vspdData.Col = C_SUB_CD
		    	.vspdData.text = arrRet(0)  
		    	.vspdData.Col = C_SUB_CD_NM
		    	.vspdData.text = arrRet(1)    
				.vspdData.action =0
        End Select
    End With
End Function
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

End Sub
'--------------------------------------------------------------------------------------------------
'	Name : openEmptName()                                                         <==== 성명/사번 팝업 
'	Description : Employee PopUp
'------------------------------------------------------------------------------------------------
Function openEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then                              'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_No.value			' Code Condition
		arrParam(1) = ""'frm1.txtName.value		    ' Name Cindition
	Else                                            'spread
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	End If
	arrParam(2) = lgUsrIntcd
	
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		      "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_Emp_No
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
'------------------------------------------------------------------------------------------------------
Function SetEmp(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then
			.txtName.value = arrRet(1)
	    	.txtEmp_no.value = arrRet(0)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(3)
			.vspdData.Col = C_DEPT_NM
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_Emp_No
			.vspdData.Text = arrRet(0)
			.vspdData.action =0	
		End If
	End With
End Function
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
	arrParam(1) = ""
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent ,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
             Case Else
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
             Case Else
        End Select
	End With
End Function       		

'========================================================================================================
' Function Name : Date_DefMask()
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Function Date_DefMask(strMaskYM)
Dim i,j
Dim ArrMask,StrComDateType
	
	Date_DefMask = False
	
	strMaskYM = ""
	
	ArrMask = Split( parent.gDateFormat, parent.gComDateType)
	
	If  parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType =  parent.gComDateType
	End If
		
	If IsArray(ArrMask) Then
		For i=0 To Ubound(ArrMask)		
			If Instr(UCase(ArrMask(i)),"D") = False Then
				If strMaskYM <> "" Then
					strMaskYM = strMaskYM & lgStrComDateType
				End If
				If Instr(UCase(ArrMask(i)),"M") And Len(ArrMask(i)) >= 3 Then
					strMaskYM = strMaskYM & "U"
					For j=0 To Len(ArrMask(i)) - 2
						strMaskYM = strMaskYM & "L"
					Next
				Else
					strMaskYM = strMaskYM & ArrMask(i)
				End If
			End If
		Next		
	Else
		Date_DefMask = False
		Exit Function
	End If	

	strMaskYM = Replace(UCase(strMaskYM),"Y","9")
	strMaskYM = Replace(UCase(strMaskYM),"M","9")

	Date_DefMask = True 
	
End Function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
    Dim strName,strDept_nm,strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	Select Case Col
         
         Case  C_EMP_NO
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_EMP_NO
    
            If Frm1.vspdData.value = "" Then
                Frm1.vspdData.Col = C_NAME
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_EMP_NO
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DEPT_NM
                Frm1.vspdData.value = ""
            Else
                IntRetCd =  FuncGetEmpInf2(Frm1.vspdData.Text,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	            If IntRetCd < 0 then
	                If  IntRetCd = -1 then
    	        		Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                    Else
                        Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                    End if
			        Frm1.vspdData.Col = C_NAME
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_DEPT_NM
                    Frm1.vspdData.value = ""
                    Set gActiveElement = document.ActiveElement
					vspdData_Change = true
                Else
                    Frm1.vspdData.Col = C_NAME
		       	    Frm1.vspdData.value = strName
		       	    Frm1.vspdData.Col = C_DEPT_NM
		       	    Frm1.vspdData.value = strDept_nm
                End if 
            End if 
  
         Case  C_SUB_TYPE_NM
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_SUB_TYPE_NM
            IntRetCd =  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_NM =  " & FilterVar(iDx , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
            If IntRetCD=False And Trim(frm1.vspdData.Text)<>""  Then
                Call  DisplayMsgBox("800142","X","X","X")      ' 코드정보에 등록되지 않은 정보입니다.
				Frm1.vspdData.Col = C_SUB_TYPE                
                frm1.vspdData.Text=""   
				vspdData_Change = true                      
            ElseIf  Parent.CountStrings(lgF0, Chr(11) ) > 1 Then      ' 같은명일 경우 pop up
                Call OpenCode(frm1.vspdData.Text, C_SUB_TYPE_POP, Row)
            ELSE
                Frm1.vspdData.Col = C_SUB_TYPE
                frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
            END IF
            
         Case  C_SUB_CD_NM  
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_SUB_CD_NM
   	        IntRetCd =  CommonQueryRs(" ALLOW_CD,ALLOW_NM "," HDA010T "," CODE_TYPE=" & FilterVar("2", "''", "S") & " AND allow_nm =  " & FilterVar(iDx , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCD=False And Trim(frm1.vspdData.Text)<>""  Then
                Call  DisplayMsgBox("800142","X","X","X")      ' 코드정보에 등록되지 않은 정보입니다.
                frm1.vspdData.Col = C_SUB_CD
                frm1.vspdData.Text=""
				vspdData_Change = true                
            ELSE
                frm1.vspdData.Col = C_SUB_CD
                frm1.vspdData.Text= Trim(Replace(lgF0,Chr(11),""))
                frm1.vspdData.Col = C_SUB_CD_NM
                frm1.vspdData.Text= Trim(Replace(lgF1,Chr(11),""))
                                                          '근태코드 
                IntRetCD =  CommonQueryRs(" ALLOW_CD,ALLOW_NM "," HDA010T "," CODE_TYPE=" & FilterVar("2", "''", "S") & " AND allow_nm =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
                If  Parent.CountStrings(lgF0, Chr(11) )>1 Then
                    Call  DisplayMsgBox("800095","X","X","X")                         '☜ : 입력된자료가 있습니다.
                    frm1.vspdData.Text=""
                Else
    	            frm1.vspdData.Col = C_SUB_CD_NM
                    frm1.vspdData.Text=Trim(frm1.vspdData.Text)
                End If
            End If 
    End Select    

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Function

'========================================================================================================
'   Event Name : txtEmp_no_Onchange           
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
			txtEmp_no_Onchange = true
            Exit function
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtsub_type_OnChange()            
'   Event Desc :
'========================================================================================================
function txtsub_type_OnChange()
    Dim iDx
    Dim IntRetCd

    IF frm1.txtsub_type.value <>"" THEN
    
        IntRetCd =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtsub_type.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        IF IntRetCd = false Then
            Call  DisplayMsgBox("800142","X","X","X")
            frm1.txtsub_type_nm.value = ""
            frm1.txtsub_type.focus
            Set gActiveElement = document.ActiveElement   
            txtsub_type_OnChange = true
            Exit function
        ELSE
            frm1.txtsub_type_nm.value = Trim(Replace(lgF0,Chr(11),""))
        END IF
        
    ELSE
        frm1.txtsub_type_nm.value = ""        
    END IF  
    
End function 
'========================================================================================================
'   Event Name : txtsub_cd_Onchange()          
'   Event Desc :
'========================================================================================================
function txtsub_cd_Onchange()
    Dim iDx
    Dim IntRetCd
    
    IF frm1.txtsub_cd.value <> "" THEN
    
        IntRetCd =  CommonQueryRs(" ALLOW_NM "," HDA010T "," CODE_TYPE=" & FilterVar("2", "''", "S") & " AND ALLOW_CD =  " & FilterVar(frm1.txtsub_cd.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        IF IntRetCd = false Then
            Call  DisplayMsgBox("800142","X","X","X")
            frm1.txtsub_cd_nm.value = ""
            frm1.txtsub_cd.focus   
			txtsub_cd_Onchange = true
            exit function
        ELSE
      
            frm1.txtsub_cd_nm.value = Trim(Replace(lgF0,Chr(11),""))

        END IF
    ELSE
        frm1.txtsub_cd_nm.value = ""
    END IF 
End function
'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If rTrim(frm1.txtFr_dept_cd.value) = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd =  FuncDeptName(frm1.txtFr_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtFr_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If rTrim(frm1.txtTo_dept_cd.value) = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd =  FuncDeptName(frm1.txtTo_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtTo_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
      gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row
     
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
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
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
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
	    Case C_NAME_POP
            Call openEmptName(1)
        Case C_EMP_NO_POP
            Call openEmptName(1)
        Case C_SUB_TYPE_POP
            Call OpenCode("", C_SUB_TYPE_POP, Row)
        Case C_SUB_CD_POP
            Call OpenCode("", C_SUB_CD_POP, Row)
    End Select 
    
End Sub
'======================================================================================================
'	Name : AutoButtonClicked()
'	Description : h4007mb2.asp 로 가는 Condition........일괄등록...........
'=======================================================================================================

Sub AutoButtonClicked(Byval ButtonDown)
	
    Dim strKeyStream
    Dim strVal
    Dim strEmp_no
    Dim strsub_type
    Dim strsub_cd
    Dim strWhere
    Dim lgStrSQL
    Dim IntRetCd
    Dim strDelete
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept, strFrYymm, strToYymm, strFrYymmSvr, strToYymmSvr
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Sub
    End If
   	strDelete = "0"

	IF frm1.txtsub_cd.value = "" then 
	    Call  DisplayMsgBox("970021","X","공제코드","X")        '공제코드를 확인하십시오.    
	    frm1.txtsub_cd.focus 
	    Exit Sub    
	End If 
    
    Fr_dept_cd = Trim(frm1.txtFr_internal_cd.value)
    To_dept_cd = Trim(frm1.txtTo_internal_cd.value)
    
    If fr_dept_cd = "" then
        IntRetCd =  FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept
		fr_dept_cd = rFrDept
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd =  FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		to_dept_cd = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
    
    If (Fr_dept_cd= "") AND (To_dept_cd="") Then       
    Else
        If Fr_dept_cd > To_dept_cd then
	        Call  DisplayMsgBox("800359","X","X","X")	'시작부서보다 작은값입니다.
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Sub
        End IF 
        
    END IF   
    
    If (frm1.txtapply_yymm_dt.Text = "") Then                       '년월의 값이 없으면 주는 기본값정의와 메시지 체크 
        strFrYymm =  UniConvYYYYMMDDToDate( parent.gDateFormat, "1900", "01", "01")
    Else
        strFrYymm =  UniConvYYYYMMDDToDate( parent.gDateFormat, frm1.txtapply_yymm_dt.Year, Right("0" & frm1.txtapply_yymm_dt.month , 2), "01")
    End if 
    
    If (frm1.txtrevoke_yymm_dt.Text = "") Then
        strToYymm =  UniConvYYYYMMDDToDate( parent.gDateFormat, "2500", "12", "31")
    Else
        strToYYMM =  UniConvYYYYMMDDToDate( parent.gDateFormat, frm1.txtrevoke_yymm_dt.Year, Right("0" & frm1.txtrevoke_yymm_dt.month , 2), "01")
    End if 
    
    
    strFrYymmSvr = UniConvDateAToB(strFrYymm, Parent.gDateFormat, Parent.gServerDateFormat)
    strToYYMMSvr = UniConvDateAToB(strToYYMM, Parent.gDateFormat, Parent.gServerDateFormat)
    If  CompareDateByFormat(strFrYymm,strToYymm,frm1.txtapply_yymm_dt.Alt,frm1.txtrevoke_yymm_dt.Alt,"970025", parent.gDateFormat, parent.gComDateType,True) = False Then
        frm1.txtapply_yymm_dt.focus
        Set gActiveElement = document.activeElement

        Exit sub
    End if 
   
    IF frm1.txtsub_cd.value <> "" THEN
        IntRetCd =  CommonQueryRs(" allow_nm "," HDA010T "," CODE_TYPE=" & FilterVar("2", "''", "S") & " AND allow_cd =  " & FilterVar(frm1.txtsub_cd.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        IF IntRetCd = false Then
            Call  DisplayMsgBox("800142","X","X","X")
            frm1.txtsub_cd.value = ""
            frm1.txtsub_cd_nm.value = ""
            frm1.txtsub_cd.focus   
        ELSE
      
            frm1.txtsub_cd_nm.value = Trim(Replace(lgF0,Chr(11),""))
        END IF
    ELSE
        frm1.txtsub_cd_nm.value = ""
    END IF 

        strEmp_no   =  FilterVar(Frm1.txtemp_no.Value & "%", "''", "S")
        strsub_type =  FilterVar(Frm1.txtsub_type.Value, "''", "S")
        strsub_cd   =  FilterVar(Frm1.txtsub_cd.Value, "''", "S")
        strWhere = " A.EMP_NO LIKE B.EMP_NO"
        strWhere = strWhere & " AND A.EMP_NO LIKE " & strEmp_no 
        strWhere = strWhere & " AND B.SUB_TYPE LIKE " & strsub_type 
        strWhere = strWhere & " AND B.SUB_CD LIKE " & strsub_cd 
        strWhere = strWhere & " AND A.PROV_TYPE = " & FilterVar("Y", "''", "S") & "  "
 ' 부서별 체크 로직 추가 - 2002.09.09 이석민 
        StrWhere = strWhere & " AND A.INTERNAL_CD between  " & FilterVar(frm1.txtFr_Internal_cd.value, "''", "S") & " AND  " & FilterVar(frm1.txtTo_Internal_cd.value, "''", "S") & ""
		strWhere = strWhere & " AND A.INTERNAL_CD LIKE  " & FilterVar(lgUsrIntcd & "%", "''", "S") & ""    '  자료권한 Internal_cd = '?'
   	   
 
   	    Call  CommonQueryRs(" COUNT(*) "," HDF020T A, HDF050T B ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
        IF Trim(Replace(lgF0,Chr(11),"")) = 0 Then
			strDelete = "Normal"
        Else
            IntRetCD = DisplayMsgBox("800502", 35,"X","X")	    '이미 생성된 자료가 있습니다.?
            If IntRetCD = vbCancel Then
	           	Exit Sub
	        
	        ELSEif IntRetCD = vbYes then
				strDelete = "Del"
			else
				strDelete = "Add"
            End If    
	    END IF
 
        frm1.vspdData.MaxRows = 0
       
       lgKeyStream  = Frm1.txtemp_No.Value & parent.gColSep                                '0
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtsub_type.Value) & parent.gColSep          '1
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtsub_cd.Value) & parent.gColSep            '2
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtFr_dept_cd.Value) & parent.gColSep        '3
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtFr_dept_nm.Value) & parent.gColSep        '4
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtto_dept_cd.Value) & parent.gColSep        '5
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtto_dept_nm.Value) & parent.gColSep        '6
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtFr_Internal_cd.Value) & parent.gColSep    '7
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtTo_Internal_cd.Value) & parent.gColSep    '8
       lgKeyStream  = lgKeyStream & lgUsrIntcd & parent.gColSep                            '9
       lgKeyStream  = lgKeyStream & strDelete & parent.gColSep

       strFrYymm =  UniConvYYYYMMDDToDate( parent.gDateFormatYYYYMM, frm1.txtapply_yymm_dt.Year,Right("0" & frm1.txtapply_yymm_dt.month , 2),"01")
       strToYymm =  UniConvYYYYMMDDToDate( parent.gDateFormatYYYYMM, frm1.txtrevoke_yymm_dt.Year,Right("0" & frm1.txtrevoke_yymm_dt.month , 2),"01")
       
       lgKeyStream  = lgKeyStream & Trim(strFrYymm) & parent.gColSep    '11
       lgKeyStream  = lgKeyStream & Trim(strToYymm) & parent.gColSep    '12       
       lgKeyStream  = lgKeyStream & Trim(Frm1.txtsub_amt.Text) & parent.gColSep    '13       

        With Frm1
    		strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                          'mb2 자동입력......						         
			strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
			strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
			strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
        End With
        
        if LayerShowHide(1) = false then
			exit Sub
		end if
        
        Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic

End Sub

'======================================================================================================
'	Name : DBAutoQueryOk()
'	Description : h5102mb2.asp 이후 Query OK해 줌 
'=======================================================================================================
Sub DBAutoQueryOk()
    Dim lRow
	Dim intIndex
	Dim daytimeVal 
	Dim strSub_type 
    
    With Frm1
        .vspdData.ReDraw = false
         ggoSpread.Source = .vspdData
   
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            .vspdData.Text =  ggoSpread.InsertFlag
        Next
            .vspdData.ReDraw = TRUE
        
    End With 
    ggoSpread.ClearSpreadData "T"
     Set gActiveElement = document.ActiveElement   
End Sub

'=======================================
'   Event Name :txtApply_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtApply_yymm_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtApply_yymm_dt.Action = 7
        frm1.txtApply_yymm_dt.focus
    End If
End Sub
'=======================================
'   Event Name : txtRevoke_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtRevoke_yymm_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtRevoke_yymm_dt.Action = 7
        frm1.txtRevoke_yymm_dt.focus
    End If
End Sub
'==========================================================================================
'   Event Name : txtApply_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtApply_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtApply_yymm_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call mainQuery()
End Sub
'==========================================================================================
'   Event Name : txtRevoke_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtRevoke_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtRevoke_yymm_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call mainQuery()
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

'========================================================================================
 ' Function Name : RemovedivTextArea
 ' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
 Function RemovedivTextArea()
 
 	Dim ii
 		
 	For ii = 1 To divTextArea.children.length
 	    divTextArea.removeChild(divTextArea.children(0))
 	Next
 
 End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>고정공제사항등록</font></td>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%>width=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
					    
						    <TR>
							   	<TD CLASS=TD5 NOWRAP>공제구분</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtsub_type" MAXLENGTH="1" SIZE=10 ALT ="공제구분" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('4')">
								                     <INPUT NAME="txtsub_type_nm" MAXLENGTH="20" SIZE=20  ALT ="Order ID" tag="14XXXU"></td>                
								<TD CLASS=TD5 NOWRAP>적용년월</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5102ma1_txtapply_yymm_dt_txtapply_yymm_dt.js'></script></TD>                 
							</TR>	
							<TR>
							    <TD CLASS=TD5 NOWRAP>공제코드</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtsub_cd"  MAXLENGTH="3" SIZE=10 ALT ="공제코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('5')">
								                     <INPUT NAME="txtsub_cd_nm" MAXLENGTH="20" SIZE=20 ALT ="Order ID" tag="14XXXU"></td>
							    <TD CLASS=TD5 NOWRAP>해제년월</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5102ma1_txtrevoke_yymm_dt_txtrevoke_yymm_dt.js'></script></TD>
							</TR>
							<TR>								
								<TD CLASS=TD5 NOWRAP>사원</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_No" MAXLENGTH="13" SIZE=10 ALT ="사번" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: openEmptName(0)">
								                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE=20 ALT ="성명" tag="14XXXU"></TD>																
    			                <TD CLASS=TD5 NOWRAP>시작부서코드</TD>
								<TD CLASS=TD6 NOWRAP>
								         <INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                             <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40 tag="14XXXU">&nbsp;~
		                                 <INPUT NAME="txtFr_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>         
							</TR> 
	                        <TR> 
								<TD CLASS=TD5 NOWRAP>공제금액</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h5102ma1_txtsub_amt_txtsub_amt.js'></script></TD>   							    
								<TD CLASS=TD5 NOWRAP>종료부서코드</TD>
								<TD CLASS=TD6 NOWRAP>
		                                 <INPUT NAME="txtto_dept_cd" MAXLENGTH="10" SIZE=10 ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							             <INPUT NAME="txtto_dept_nm" MAXLENGTH="40" SIZE=20 ALT ="Order ID" tag="14XXXU">
    			                         <INPUT NAME="txtTo_Internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7 tag="14XXXU"></TD>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h5102ma1_vaSpread_vspdData.js'></script>
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
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD width=100><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: AutoButtonClicked('1')">자동입력</BUTTON>&nbsp;</TD>
	                <TD Width=*>&nbsp;</TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>   
	
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=Bizsize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>

</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
