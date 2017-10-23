<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Single
*  3. Program ID           : h6027ma1
*  4. Program Name         : h6027ma1
*  5. Program Desc         : 원천징수이행상황신고 
*  6. Comproxy List        :
*  7. Modified date(First) : 2004/03/04
*  8. Modified date(Last)  : 2004/03/04
*  9. Modifier (First)     : 최용철 
* 10. Modifier (Last)      : 최용철 
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

<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID       = "h6027mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID2      = "h6027mb2.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim Biz_LogicCheckFlag
Dim FlgModeCheck

Dim RevertYYMM_Default
Dim RevertYYMM_Change
Dim ProvYYMM_Default
Dim ProvYYMM_Change
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
	lgOldRow = 0
		
    gblnWinEvent      = False
	lgBlnFlawChgFlg   = False

    Biz_LogicCheckFlag = False
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    lgKeyStream  = ""
    lgKeyStream  = lgKeyStream & Trim(Frm1.txtComp_cd.value) & parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtprov_yymm.Year & Right("0" & Frm1.txtprov_yymm.Month,2) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtrevert_yymm.Year & Right("0" & Frm1.txtrevert_yymm.Month,2) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtsubmit_yymm.Year & Right("0" & Frm1.txtsubmit_yymm.Month, 2) & Right("0" & Frm1.txtsubmit_yymm.Day, 2) & Parent.gColSep  
    lgKeyStream  = lgKeyStream & lgUsrIntcd & parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtRetireFr_dt.Year & Right("0" & Frm1.txtRetireFr_dt.Month, 2) & Right("0" & Frm1.txtRetireFr_dt.Day, 2) & Parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtRetireTo_dt.Year & Right("0" & Frm1.txtRetireTo_dt.Month, 2) & Right("0" & Frm1.txtRetireTo_dt.Day, 2) & Parent.gColSep 
    lgKeyStream  = lgKeyStream & Frm1.txtYearEnd_yymm.Year & Right("0" & Frm1.txtYearEnd_yymm.Month,2) & Parent.gColSep
    
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

    '신고 사업장    
    Call CommonQueryRs("YEAR_AREA_NM,YEAR_AREA_CD","HFA100T","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0
    iCodeArr = lgF1   
    Call SetCombo2(frm1.txtComp_cd,iCodeArr,iNameArr,Chr(11))   

End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtprov_yymm.focus()

	Call ExtractDateFrom("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
	Call ggoOper.FormatDate(frm1.txtprov_yymm, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtrevert_yymm, parent.gDateFormat, 2)		
	Call ggoOper.FormatDate(frm1.txtsubmit_yymm, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtRetireFr_dt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtRetireTo_dt, parent.gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtYearEnd_yymm, parent.gDateFormat, 2)
		
	frm1.txtprov_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtprov_yymm.Month = strMonth

	frm1.txtrevert_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtrevert_yymm.Month = strMonth

	frm1.txtsubmit_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtsubmit_yymm.Month = strMonth
	frm1.txtsubmit_yymm.Day   = "10"
	strDay = "10"
			
	frm1.txtRetireFr_dt.Year = strYear 	 
	frm1.txtRetireFr_dt.Month = strMonth
	frm1.txtRetireFr_dt.Day   = strDay
	
	frm1.txtRetireTo_dt.Year = strYear 	 
	frm1.txtRetireTo_dt.Month = strMonth
	frm1.txtRetireTo_dt.Day   = strDay
	
	frm1.txtYearEnd_yymm.Year = strYear 		 '년월일 default value setting
	frm1.txtYearEnd_yymm.Month = strMonth
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Dim strYear
    Dim strMonth
    Dim strDay

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
	Call AppendNumberPlace("6","5","0")
    Call ggoOper.LockField(Document, "N")                     '⊙: Lock Field

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call InitVariables                                                              'Initializes local global variables    
	Call InitComboBox
    Call SetDefaultVal
    
	Call SetToolbar("11001000000001")												'⊙: Set ToolBar

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    
    If lgBlnFlgChgValue = True AND Biz_LogicCheckFlag = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call SetToolbar("11001000000001")

 	      
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                         '⊙: Initializes local global variables
 
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field       
    Call MakeKeyStream("Q")
    Call DisableToolBar(Parent.TBC_QUERY)

	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If


    FncQuery = True                                                              '☜: Processing is OK
'    Biz_LogicCheckFlag = True

End Function

'========================================================================================================
' Name : FncQuery2
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery2()
    Dim IntRetCD 
    
    FncQuery2 = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call SetToolbar("11001000000001")
 
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                         '⊙: Initializes local global variables
 
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field       
    Call MakeKeyStream("Q")
    Call DisableToolBar(Parent.TBC_QUERY)

	If DBQuery2 = False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If


    FncQuery2 = True                                                              '☜: Processing is OK
'    Biz_LogicCheckFlag = True

End Function
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 

    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If

    Call ggoOper.ClearField(Document, "2")                                       '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    Call SetToolbar("11001000000001")
    Call SetDefaultVal
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   

    Biz_LogicCheckFlag = True
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If


    Call DataChecking()
    If FlgModeCheck > 0 and ((ProvYYMM_Default <> ProvYYMM_Change) OR (RevertYYMM_Default <> RevertYYMM_Change)) Then
		IntRetCD = DisplayMsgBox("800604", "x" , "조회조건" ,"x")
		If IntRetCD <> vbYes then
			Call FncQuery()
			Exit Function
		Else
			Exit Function
		End If
    End If
        
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
   
    Call MakeKeyStream("D")
    Call  DisableToolBar( parent.TBC_DELETE)
    If DbDelete = False Then
		Exit Function
	End If

    Set gActiveElement = document.ActiveElement   
    
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

'    If lgBlnFlgChgValue = False Then 
 '       IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
  '      Exit Function
   ' End If

    Call DataChecking()
    If FlgModeCheck > 0 and ((ProvYYMM_Default <> ProvYYMM_Change) OR (RevertYYMM_Default <> RevertYYMM_Change)) Then
		IntRetCD = DisplayMsgBox("800604", "x" , "조회조건" ,"x")
		If IntRetCD <> vbYes then
			Call FncQuery()
			Exit Function
		Else
			Exit Function
		End If
    End If


    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

    Call MakeKeyStream("S")

    Call  DisableToolBar( parent.TBC_SAVE)
    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                      '☜: Protect system from crashing
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
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbQuery2
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery2()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery2 = False                                                              '☜: Processing is NG

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID2 & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal      & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal      & "&txtPrevNext="      & ""	                             '☜: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    
    DbQuery2 = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Name : DataChecking
' Desc : DataChecking Of Save/Delete/AutoQuery
'========================================================================================================
Function DataChecking()
    Dim strWhere    

    strWhere = " biz_area_cd = " & FilterVar(Frm1.txtComp_cd.value, "''", "S") & ""
    strWhere = strWhere & " AND prov_yymm =  " & FilterVar(ProvYYMM_Default , "''", "S") & ""      

    Call  CommonQueryRs(" COUNT(*) "," HDF500T ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	FlgModeCheck = Trim(Replace(lgF0,Chr(11),""))

    ProvYYMM_Change   = Frm1.txtprov_yymm.Year & Right("0" & Frm1.txtprov_yymm.Month,2)
    RevertYYMM_Change = Frm1.txtrevert_yymm.Year & Right("0" & Frm1.txtrevert_yymm.Month,2)
End Function

'======================================================================================================
'	Name : AutoButtonClicked()
'	Description : h6027mb2.asp 로 가는 Condition........일괄등록...........
'=======================================================================================================
Sub AutoButtonClicked(Byval ButtonDown)
	Dim IntRetCD
	
	Call DataChecking()  

    IF FlgModeCheck = 0 Then
        Call FncQuery2()
    Else
        IntRetCD = DisplayMsgBox("800602", parent.VB_YES_NO , RevertYYMM_Default ,"귀속년월 기준")	'%1월 %2(으)로 이미 생성되어 있습니다. 다시 생성하시겠습니까?

        If IntRetCD = vbYes then
            Call FncQuery2()
        Else
            Call FncQuery()
        End If    
    END IF
End Sub

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
	Call LayerShowHide(1)
		
	With Frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	Call LayerShowHide(1)
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strVal

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = false

	Call SetToolbar("11011000000001")
	
	
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

    ProvYYMM_Default   = Frm1.txtprov_yymm.Year & Right("0" & Frm1.txtprov_yymm.Month,2)
    RevertYYMM_Default = Frm1.txtrevert_yymm.Year & Right("0" & Frm1.txtrevert_yymm.Month,2)

    Biz_LogicCheckFlag = True
	
End Function

'========================================================================================================
' Function Name : DbQueryNG
' Function Desc : Called by MB Area when query operation is fail
'========================================================================================================
Function DbQueryNG()
    Dim strVal

    lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = false

    Call SetToolbar("11001000000001")
	
	
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

    ProvYYMM_Default   = Frm1.txtprov_yymm.Year & Right("0" & Frm1.txtprov_yymm.Month,2)
    RevertYYMM_Default = Frm1.txtrevert_yymm.Year & Right("0" & Frm1.txtrevert_yymm.Month,2)

    Biz_LogicCheckFlag = True

    Call ggoOper.SetReqAttr(frm1.txtrevert_yymm, "N")
    Call ggoOper.SetReqAttr(frm1.txtsubmit_yymm, "N")
    Call ggoOper.SetReqAttr(frm1.txtRetireFr_dt, "N")
    Call ggoOper.SetReqAttr(frm1.txtRetireTo_dt, "N")
    Call ggoOper.SetReqAttr(frm1.txtYearEnd_yymm, "N")
End Function
'========================================================================================================
' Function Name : DbQueryOk2
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk2()
    Dim strVal

    If FlgModeCheck = 0 Then
		lgIntFlgMode      = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
	Else
		lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    End If
    lgBlnFlgChgValue = false

	Call SetToolbar("11001000000001")

    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

    lgBlnFlgChgValue   = True
    Biz_LogicCheckFlag = True
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables	
    Call FncQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
    Call InitVariables()
    Call FncQuery()
End Function

'========================================================================================
' Function Name : btnPreview_onClick()
' Function Desc : PREVIEW
'========================================================================================
Function FncBtnPreview()
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    Dim strUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile , StrEbrFile2 , StrEbrFile3, StrEbrFile4
    Dim ObjName	, prov_yymm , revert_yymm, submit_yymm , Comp_cd , IntRetCD

	
'    If lgIntFlgMode = Parent.OPMD_CMODE Then
'		Call DisplayMsgBox("800167","X","X","X")
'		Exit Function
 '   Else
'		If lgBlnFlgChgValue = True Then
'		    IntRetCD =  DisplayMsgBox("900017",parent.VB_YES_NO ,"X","X")  '데이터가 변경되었습니다. 계속하시겠습니까?
'			If IntRetCD = vbNo Then
'	    		Exit Function
'			End If
 '       End If		
  '  End If
		
    StrEbrFile  = "h6027oa1"
    StrEbrFile2 = "h6027oa1_2"
    StrEbrFile3 = "h6027oa1_3"
    StrEbrFile4 = "h6027oa1_4"
	
    prov_yymm   = frm1.txtprov_yymm.Year & Right("0" & frm1.txtprov_yymm.Month, 2)
    revert_yymm = frm1.txtrevert_yymm.Year & Right("0" & frm1.txtrevert_yymm.Month, 2)
    submit_yymm = frm1.txtsubmit_yymm.Year & Right("0" & frm1.txtsubmit_yymm.Month, 2) & Right("0" & frm1.txtsubmit_yymm.Day, 2) & Parent.gColSep  
    Comp_cd = Trim(frm1.txtComp_cd.value)
    
    strUrl = "prov_yymm|" & prov_yymm
    strUrl = strUrl & "|Comp_cd|" & Comp_cd


    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
    Call FncEBRPreview(ObjName, strUrl)
		
    ObjName = AskEBDocumentName(StrEbrFile3,"ebr")
    Call FncEBRPreview(ObjName, strUrl)
	
    If frm1.prt_check1_flag.checked = True Then
   	ObjName = AskEBDocumentName(StrEbrFile2,"ebr")
	Call FncEBRPreview(ObjName, strUrl)

	ObjName = AskEBDocumentName(StrEbrFile4,"ebr")
        Call FncEBRPreview(ObjName, strUrl)
    End If
			
End Function

'========================================================================================
' Function Name : FncBtnPrint_onClick()
' Function Desc : PRIENT
'========================================================================================
Function FncBtnPrint()
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    Dim strUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile , StrEbrFile2 , StrEbrFile3, StrEbrFile4
    Dim ObjName	, prov_yymm , revert_yymm , submit_yymm , Comp_cd , IntRetCD


'    If lgIntFlgMode = Parent.OPMD_CMODE Then
'	Call DisplayMsgBox("800167","X","X","X")
'	Exit Function
 '   Else
'		If lgBlnFlgChgValue = True Then
'		    IntRetCD =  DisplayMsgBox("900017",parent.VB_YES_NO ,"X","X")  '데이터가 변경되었습니다. 계속하시겠습니까?
'			If IntRetCD = vbNo Then
'	    		Exit Function
'			End If
 '       End If		
  '  End If

    StrEbrFile  = "h6027oa1"
    StrEbrFile2 = "h6027oa1_2"
    StrEbrFile3 = "h6027oa1_3"
    StrEbrFile4 = "h6027oa1_4"
	
    prov_yymm   = frm1.txtprov_yymm.Year & Right("0" & frm1.txtprov_yymm.Month, 2)
    revert_yymm = frm1.txtrevert_yymm.Year & Right("0" & frm1.txtrevert_yymm.Month, 2)
    submit_yymm = frm1.txtsubmit_yymm.Year & Right("0" & frm1.txtsubmit_yymm.Month, 2) & Right("0" & frm1.txtsubmit_yymm.Day, 2) & Parent.gColSep  
    
    Comp_cd = Trim(frm1.txtComp_cd.value)
    
    strUrl = "prov_yymm|" & prov_yymm
    strUrl = strUrl & "|Comp_cd|" & Comp_cd

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
    Call FncEBRPrint(EBAction , ObjName , strUrl)
		
    ObjName = AskEBDocumentName(StrEbrFile3,"ebr")
    Call FncEBRPrint(EBAction , ObjName , strUrl)
			
    If frm1.prt_check1_flag.checked = True Then
   	ObjName = AskEBDocumentName(StrEbrFile2,"ebr")
	Call FncEBRPrint(EBAction , ObjName , strUrl)
	ObjName = AskEBDocumentName(StrEbrFile4,"ebr")
	Call FncEBRPrint(EBAction , ObjName , strUrl)	
    End If

End Function
'******************************  3.2.1 Object TSub txt_i_Ag 처리  ***********************************
'	Window에 발생 하는 모든 Even 처리	
'****************************************************************************************************
Dim T_SUM

Dim T_013, T_023, T_033, T_043
Dim T_014, T_024, T_044, T_264
Dim T_015, T_025, T_035, T_045
Dim T_253, T_263
Dim T_255, T_265
Dim T_103, T_203, T_303, T_403, T_453, T_503, T_603, T_693, T_803, T_903
Dim T_104, T_304, T_504, T_604, T_694, T_904
Dim T_105, T_205, T_305, T_405, T_455, T_505, T_605, T_695, T_805, T_905
Dim T_106, T_206, T_306, T_406, T_456, T_506, T_606, T_696, T_806, T_906

Function Chk_Option(ByRef prObject)
	Dim vValue 
	vValue = prObject.Value

	If UCase(vValue) = "T" or vValue = "△" Then
		prObject.value = "△"
	ElseIf UCase(vValue) = "S" or vValue = "□" Then
		prObject.value = "□"
	Else
		prObject.value = ""
	End If
End Function

	    
Function Biz_Logic(Byval value)
	With frm1
		Select Case value
			'****A10 가감계****
			Case "A101"
			    .txt_i_A101.text = 0
			    .txt_i_A101.text = UNICDbl(.txt_i_A011.text) + UNICDbl(.txt_i_A021.text) + UNICDbl(.txt_i_A031.text) + UNICDbl(.txt_i_A041.text)

			Case "A102"
			    .txt_i_A102.text = 0
 			    .txt_i_A102.text = UNICDbl(.txt_i_A012.text) + UNICDbl(.txt_i_A022.text) + UNICDbl(.txt_i_A032.text) + UNICDbl(.txt_i_A042.text)

			Case "A103"
				T_SUM = 0
			    .txt_i_A103.text = 0

                T_013 = UNICDbl(.txt_i_A013.text)
                T_023 = UNICDbl(.txt_i_A023.text)
                T_033 = UNICDbl(.txt_i_A033.text)
                T_043 = UNICDbl(.txt_i_A043.text)

				T_SUM = T_013 + T_023 + T_033 + T_043

	           .txt_i_A103.text = T_SUM 
                Call Biz_Logic2("A107")   
                
			Case "A104"

				T_SUM = 0
			    .txt_i_A104.text = 0

                T_014 = UNICDbl(.txt_i_A014.text)
                T_024 = UNICDbl(.txt_i_A024.text)
                T_044 = UNICDbl(.txt_i_A044.text)

				T_SUM = T_014 + T_024 + T_044

		        .txt_i_A104.text = T_SUM

			Case "A105"
				T_SUM = 0
			    .txt_i_A105.text = 0

                 T_015 = UNICDbl(.txt_i_A015.text)
                 T_025 = UNICDbl(.txt_i_A025.text)
                 T_035 = UNICDbl(.txt_i_A035.text)
                 T_045 = UNICDbl(.txt_i_A045.text)

				T_SUM = T_015 + T_025 + T_035 + T_045

	            .txt_i_A105.text = T_SUM                
                
			'****A30 가감계****
			Case "A301"
			    .txt_i_A301.text = 0
			    .txt_i_A301.text = UNICDbl(.txt_i_A251.text) + UNICDbl(.txt_i_A261.text)
			Case "A302"
			    .txt_i_A302.text = 0
			    .txt_i_A302.text = UNICDbl(.txt_i_A252.text) + UNICDbl(.txt_i_A262.text)
			Case "A303"
				T_SUM = 0
			    .txt_i_A303.text = 0

                T_253 = UNICDbl(.txt_i_A253.text)
                T_263 = UNICDbl(.txt_i_A263.text)

				T_SUM = T_253 + T_263
	            .txt_i_A303.text = T_SUM
                
			Case "A304"
				T_SUM = 0
			    .txt_i_A304.text = 0

                T_264 = UNICDbl(.txt_i_A264.text)
				T_SUM = T_264

	            .txt_i_A304.text = T_SUM
	            
			Case "A305"
				T_SUM = 0
			    .txt_i_A305.text = 0

                T_255 = UNICDbl(.txt_i_A255.text)
                T_265 = UNICDbl(.txt_i_A265.text)

				T_SUM = T_255 + T_265
	            .txt_i_A305.text = T_SUM			

			'****A90 가감계****
			Case "A991"
			    .txt_i_A991.text = 0
			    .txt_i_A991.text = UNICDbl(.txt_i_A101.text) + UNICDbl(.txt_i_A201.text) + UNICDbl(.txt_i_A301.text) + UNICDbl(.txt_i_A401.text) +_
			                       UNICDbl(.txt_i_A451.text) + UNICDbl(.txt_i_A501.text) + UNICDbl(.txt_i_A601.text) + UNICDbl(.txt_i_A691.text) +_
			                       UNICDbl(.txt_i_A801.text)
			Case "A992"
			    .txt_i_A992.text = 0
			    .txt_i_A992.text = UNICDbl(.txt_i_A102.text) + UNICDbl(.txt_i_A202.text) + UNICDbl(.txt_i_A302.text) + UNICDbl(.txt_i_A402.text) +_
			                       UNICDbl(.txt_i_A452.text) + UNICDbl(.txt_i_A502.text) + UNICDbl(.txt_i_A602.text) + UNICDbl(.txt_i_A802.text)
			Case "A993"

				T_SUM = 0
			    .txt_i_A993.text = 0
	
				T_103 = UNICDbl(.txt_i_A103.text)
				T_203 = UNICDbl(.txt_i_A203.text)
				T_303 = UNICDbl(.txt_i_A303.text)
				T_403 = UNICDbl(.txt_i_A403.text)
				T_453 = UNICDbl(.txt_i_A453.text)
				T_503 = UNICDbl(.txt_i_A503.text)
				T_603 = UNICDbl(.txt_i_A603.text)
				T_693 = UNICDbl(.txt_i_A693.text)
				T_803 = UNICDbl(.txt_i_A803.text)
				T_903 = UNICDbl(.txt_i_A903.text)

				T_SUM = T_103 + T_203 + T_303 + T_403 + T_453 + T_503 + T_603 + T_693 + T_803 + T_903
	            .txt_i_A993.text = T_SUM
                
			Case "A994"
				T_SUM = 0
			    .txt_i_A994.text = 0
				T_104 = UNICDbl(.txt_i_A104.text)
				T_304 = UNICDbl(.txt_i_A304.text)
				T_504 = UNICDbl(.txt_i_A504.text)
				T_604 = UNICDbl(.txt_i_A604.text)
				T_694 = UNICDbl(.txt_i_A694.text)
				T_904 = UNICDbl(.txt_i_A904.text)

				T_SUM = T_104 + T_304 + T_504 + T_604 +T_694 + T_904
	            .txt_i_A994.text = T_SUM 


			Case "A995"
				T_SUM = 0

			    .txt_i_A995.text = 0


				T_105 = UNICDbl(.txt_i_A105.text)
				T_205 = UNICDbl(.txt_i_A205.text)
				T_305 = UNICDbl(.txt_i_A305.text)
				T_405 = UNICDbl(.txt_i_A405.text)
				T_455 = UNICDbl(.txt_i_A455.text)
				T_505 = UNICDbl(.txt_i_A505.text)
				T_605 = UNICDbl(.txt_i_A605.text)
				T_695 = UNICDbl(.txt_i_A695.text)
				T_805 = UNICDbl(.txt_i_A805.text)
				T_905 = UNICDbl(.txt_i_A905.text)

				T_SUM = T_105 + T_205 + T_305 + T_405 + T_455 + T_505 + T_605 + T_695 + T_805 + T_905

		        .txt_i_A995.text = T_SUM
              
                
			Case "A996"
				T_SUM = 0
			    .txt_i_A996.text = 0

				T_106 = UNICDbl(.txt_i_A106.text)
				T_206 = UNICDbl(.txt_i_A206.text)
				T_306 = UNICDbl(.txt_i_A306.text)
				T_406 = UNICDbl(.txt_i_A406.text)
				T_456 = UNICDbl(.txt_i_A456.text)
				T_506 = UNICDbl(.txt_i_A506.text)
				T_606 = UNICDbl(.txt_i_A606.text)
				T_696 = UNICDbl(.txt_i_A696.text)
				T_806 = UNICDbl(.txt_i_A806.text)
				T_906 = UNICDbl(.txt_i_A906.text)

				T_SUM = T_106 + T_206 + T_306 + T_406 + T_456 + T_506 + T_606 + T_696 + T_806 + T_906

		        .txt_i_A996.text = T_SUM

			Case "A997"
			    .txt_i_A997.text = 0
			    .txt_i_A997.text = UNICDbl(.txt_i_A107.text) + UNICDbl(.txt_i_A207.text) + UNICDbl(.txt_i_A307.text) + UNICDbl(.txt_i_A407.text) +_
			                       UNICDbl(.txt_i_A457.text) + UNICDbl(.txt_i_A507.text) + UNICDbl(.txt_i_A607.text) + UNICDbl(.txt_i_A697.text) +_
			                       UNICDbl(.txt_i_A807.text) + UNICDbl(.txt_i_A907.text)


			Case "A998"
			    .txt_i_A998.text = 0
			    .txt_i_A998.text = UNICDbl(.txt_i_A108.text) + UNICDbl(.txt_i_A308.text) + UNICDbl(.txt_i_A508.text) + UNICDbl(.txt_i_A608.text) +_
			                       UNICDbl(.txt_i_A698.text) + UNICDbl(.txt_i_A908.text)
	    End Select
	    
	    
	End With
End function

' 2005/08/04	납부세액 부분 로직 추가 (소득세 + 가산세)     
Function Biz_Logic2(Byval value)
	
	With frm1
	
		Select Case value
		
			Case "A107"
				.txt_i_A107.text = 0
				.txt_i_A107.text = UNICDbl(.txt_i_A103.text) + UNICDbl(.txt_i_A105.text)
			
			Case "A207"
				.txt_i_A207.text = 0
				.txt_i_A207.text = UNICDbl(.txt_i_A203.text) + UNICDbl(.txt_i_A205.text)

			Case "A407"
				.txt_i_A407.text = 0
				.txt_i_A407.text = UNICDbl(.txt_i_A403.text) + UNICDbl(.txt_i_A405.text)	

			Case "A457"
				.txt_i_A457.text = 0
				.txt_i_A457.text = UNICDbl(.txt_i_A453.text) + UNICDbl(.txt_i_A455.text)	
								
			Case "A507"
				.txt_i_A507.text = 0
				.txt_i_A507.text = UNICDbl(.txt_i_A503.text) + UNICDbl(.txt_i_A505.text)	
				
			Case "A607"
				.txt_i_A607.text = 0
				.txt_i_A607.text = UNICDbl(.txt_i_A603.text) + UNICDbl(.txt_i_A605.text)
				
			Case "A697"
				.txt_i_A697.text = 0
				.txt_i_A697.text = UNICDbl(.txt_i_A693.text) + UNICDbl(.txt_i_A695.text)		
				
			Case "A807"
				.txt_i_A807.text = 0
				.txt_i_A807.text = UNICDbl(.txt_i_A803.text) + UNICDbl(.txt_i_A805.text)	
				
			Case "A907"
				.txt_i_A907.text = 0
				.txt_i_A907.text = UNICDbl(.txt_i_A903.text) + UNICDbl(.txt_i_A905.text)										
		 End Select	    
	    
	End With
End function


'===================================
'1.인원 
'===================================
Sub txt_i_A011_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A101")
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A021_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A101")
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A031_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A101")
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A041_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A101")
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A101_Change()
	lgBlnFlgChgValue = True
    '--> Call Biz_Logic(A991)
End Sub
Sub txt_i_A201_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A251_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A301")
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A261_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A301")
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A301_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A401_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A451_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A501_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A601_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A691_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A801_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A991")
    End If
End Sub
Sub txt_i_A991_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub


'===================================
'2.총지급액 
'===================================
Sub txt_i_A012_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A102")
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A022_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A102")
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A032_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A102")
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A042_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A102")
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A102_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A202_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A252_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A302")
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A262_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A302")
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A302_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A402_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A452_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A502_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A602_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A802_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A992")
    End If
End Sub
Sub txt_i_A992_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub



'===================================
'3.소득세등 
'===================================
Sub txt_i_A013_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A103")
	    Call Biz_Logic("A993")
    End If
End Sub
Sub txt_i_A023_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A103")
	    Call Biz_Logic("A993")
    End If
End Sub
Sub txt_i_A033_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A103")
	    Call Biz_Logic("A993")
    End If
End Sub
Sub txt_i_A043_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A103")
	    Call Biz_Logic("A993")
    End If
End Sub
Sub txt_i_A203_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
		Call Biz_Logic2("A207")	    
    End If
End Sub
Sub txt_i_A253_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A303")
	    Call Biz_Logic("A993")
    End If
End Sub
Sub txt_i_A263_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A303")
	    Call Biz_Logic("A993")
    End If
End Sub
Sub txt_i_A303_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A403_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
	    Call Biz_Logic2("A407")	
    End If
End Sub
Sub txt_i_A453_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
	    Call Biz_Logic2("A457")	
    End If
End Sub
Sub txt_i_A503_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
	    Call Biz_Logic2("A507")	
    End If
End Sub
Sub txt_i_A603_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
	    Call Biz_Logic2("A607")	
    End If
End Sub
Sub txt_i_A693_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
	    Call Biz_Logic2("A697")	
    End If
End Sub
Sub txt_i_A803_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
	    Call Biz_Logic2("A807")	
    End If
End Sub
Sub txt_i_A903_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A993")
	    Call Biz_Logic2("A907")	
    End If
End Sub
Sub txt_i_A993_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub


'===================================
'4.농어촌특별세 
'===================================
Sub Chk_i_A014_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A014)
		Call Biz_Logic("A104")
	    Call Biz_Logic("A994")
    End If
End Sub
Sub Chk_i_A024_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A024)
		Call Biz_Logic("A104")
	    Call Biz_Logic("A994")
    End If
End Sub
Sub Chk_i_A044_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A044)
		Call Biz_Logic("A104")
	    Call Biz_Logic("A994")
    End If
End Sub

Sub Chk_i_A264_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A264)
		Call Biz_Logic("A304")
	    Call Biz_Logic("A994")
    End If
End Sub

Sub Chk_i_A504_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A504)
	    Call Biz_Logic("A994")
    End If
End Sub
Sub Chk_i_A604_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A604)
	    Call Biz_Logic("A994")
    End If
End Sub
Sub Chk_i_A694_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A694)
	    Call Biz_Logic("A994")
    End If
End Sub
Sub Chk_i_A904_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A904)
	    Call Biz_Logic("A994")
    End If
End Sub



Sub txt_i_A014_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A104")
	    Call Biz_Logic("A994")
    End If
End Sub
Sub txt_i_A024_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A104")
	    Call Biz_Logic("A994")
    End If
End Sub
Sub txt_i_A044_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A104")
	    Call Biz_Logic("A994")
    End If
End Sub
Sub txt_i_A104_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A264_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A304")
	    Call Biz_Logic("A994")
    End If
End Sub
Sub txt_i_A304_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A504_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A994")
	    frm1.txt_i_A508.value = frm1.txt_i_A504.value
    End If
End Sub
Sub txt_i_A604_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A994")
	    frm1.txt_i_A608.value = frm1.txt_i_A604.value
    End If
End Sub
Sub txt_i_A694_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A994")
	    frm1.txt_i_A698.value = frm1.txt_i_A694.value
    End If
End Sub
Sub txt_i_A904_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A994")
	    frm1.txt_i_A908.value = frm1.txt_i_A904.value
    End If
End Sub
Sub txt_i_A994_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub


'===================================
'5.가산세 
'===================================
Sub Chk_i_A015_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A015)
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A025_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A025)
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A035_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A035)
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A045_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A045)
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub

Sub Chk_i_A205_OnChange()
msgbox 1
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A205)
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A255_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A255)
		Call Biz_Logic("A305")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A265_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A265)
		Call Biz_Logic("A305")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A405_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A405)
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A455_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A455)
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A505_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A505)
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A605_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A605)
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A695_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A695)
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A805_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A805)
	    Call Biz_Logic("A995")
    End If
End Sub
Sub Chk_i_A905_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A905)
	    Call Biz_Logic("A995")
    End If
End Sub




Sub txt_i_A015_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub txt_i_A025_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub txt_i_A035_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub txt_i_A045_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A105")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub txt_i_A105_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A205_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
		Call Biz_Logic2("A207")	    
    End If
End Sub
Sub txt_i_A255_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A305")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub txt_i_A265_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
		Call Biz_Logic("A305")
	    Call Biz_Logic("A995")
    End If
End Sub
Sub txt_i_A305_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
Sub txt_i_A405_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
	    Call Biz_Logic2("A407")	
    End If
End Sub
Sub txt_i_A455_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
	    Call Biz_Logic2("A457")	
    End If
End Sub
Sub txt_i_A505_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
	    Call Biz_Logic2("A507")	
    End If
End Sub
Sub txt_i_A605_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
	    Call Biz_Logic2("A607")	
    End If
End Sub
Sub txt_i_A695_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
	    Call Biz_Logic2("A697")	
    End If
End Sub
Sub txt_i_A805_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
	    Call Biz_Logic2("A807")	
    End If
End Sub
Sub txt_i_A905_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A995")
	    Call Biz_Logic2("A907")	
    End If
End Sub
Sub txt_i_A995_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub


'===================================
'6.당월조정환급세액 
'===================================
Sub Chk_i_A106_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A106)
	    Call Biz_Logic("A996")
    End If
End Sub

Sub Chk_i_A206_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A206)
	    Call Biz_Logic("A996")
    End If
End Sub

Sub Chk_i_A306_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A306)
	    Call Biz_Logic("A996")
    End If
End Sub

Sub Chk_i_A406_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A406)
	    Call Biz_Logic("A996")
    End If
End Sub
Sub Chk_i_A456_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A456)
	    Call Biz_Logic("A996")
    End If
End Sub
Sub Chk_i_A506_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A506)
	    Call Biz_Logic("A996")
    End If
End Sub
Sub Chk_i_A606_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A606)
	    Call Biz_Logic("A996")
    End If
End Sub
Sub Chk_i_A696_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A696)
	    Call Biz_Logic("A996")
    End If
End Sub
Sub Chk_i_A806_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A806)
	    Call Biz_Logic("A996")
    End If
End Sub
Sub Chk_i_A906_OnChange()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Chk_Option(frm1.Chk_i_A906)
	    Call Biz_Logic("A996")
    End If
End Sub




Sub txt_i_A106_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A206_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A306_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A406_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A456_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A506_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A606_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A696_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A806_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A906_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A996")
    End If
End Sub
Sub txt_i_A996_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub
'===================================
'7.소득세등(가산세포함)
'===================================
Sub txt_i_A107_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A207_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A307_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A407_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A457_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A507_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A607_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A697_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A807_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub
Sub txt_i_A907_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A997")
    End If
End Sub

Sub txt_i_A997_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub


'===================================
'8.농어촌특별세 
'===================================
Sub txt_i_A108_Change()
	lgBlnFlgChgValue = True
    Call Biz_Logic("A998")
End Sub
Sub txt_i_A308_Change()
	lgBlnFlgChgValue = True
    Call Biz_Logic("A998")
End Sub
Sub txt_i_A508_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A998")
    End If
End Sub
Sub txt_i_A608_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A998")
    End If
End Sub
Sub txt_i_A698_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A998")
    End If
End Sub
Sub txt_i_A908_Change()
	lgBlnFlgChgValue = True
    If Biz_LogicCheckFlag = True Then
	    Call Biz_Logic("A998")
    End If
End Sub
Sub txt_i_A998_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub



'===================================
'II.환급세액조정 
'===================================
Sub txt_ii_A001_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A002_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A003_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A004_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A005_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A006_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A007_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A008_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

Sub txt_ii_A009_Change()
	lgBlnFlgChgValue = True
'    '--> Call Biz_Logic(A1)
End Sub

'-------------------------------
Function txtrevert_yymm_Change()
	lgBlnFlgChgValue = True

End Function 

Function txtsubmit_yymm_Change()
	lgBlnFlgChgValue = True

End Function 
Function txtRetireFr_dt_Change()
	lgBlnFlgChgValue = True
End Function 
Function txtRetireTo_dt_Change()
	lgBlnFlgChgValue = True
End Function 

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtprov_yymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtprov_yymm.Action = 7
        frm1.txtprov_yymm.focus
    End If
End Sub
Sub txtrevert_yymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtrevert_yymm.Action = 7
        frm1.txtrevert_yymm.focus
    End If
End Sub
Sub txtsubmit_yymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtsubmit_yymm.Action = 7
        frm1.txtsubmit_yymm.focus
    End If
End Sub
Sub txtRetireFr_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtRetireFr_dt.Action = 7
        frm1.txtRetireFr_dt.focus
    End If
End Sub
Sub txtRetireFr_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtRetireFr_dt.Action = 7
        frm1.txtRetireFr_dt.focus
    End If
End Sub
Sub txtYearEnd_yymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtprov_yymm.Action = 7
        frm1.txtprov_yymm.focus
    End If
End Sub

Sub txtprov_yymm_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery   'Call FncQuery()
End Sub
Sub txtrevert_yymm_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery   'Call FncQuery()
End Sub
Sub txtsubmit_yymm_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery   'Call FncQuery()
End Sub



</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>


<BODY TABINDEX="-1" SCROLL="YES">
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>원천징수이행상황신고</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></TD>
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
	   	<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%>></TD>
			    </TR>
				<TR>
					<TD HEIGHT="20" WIDTH="100%">
					    <FIELDSET CLASS="CLSFLD">
					        <TABLE <%=LR_SPACE_TYPE_40%>>
					 	        <TR>
									<TD CLASS=TD5 NOWRAP>지급연월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtprov_yymm NAME="txtprov_yymm" CLASS=FPDTYYYYMM title=FPDATETIME ALT="지급연월" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>		
							        <TD CLASS=TD5 NOWRAP>사업장</TD>
  								    <TD CLASS="TD6" NOWRAP><SELECT NAME="txtComp_cd" ALT="신고사업장" STYLE="WIDTH: 150px" TAG="12N"></SELECT></TD>
  								    <TD><BUTTON NAME="btnPreview"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON><BUTTON NAME="btnPrint"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>귀속연월</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtrevert_yymm NAME="txtrevert_yymm" CLASS=FPDTYYYYMM title=FPDATETIME ALT="귀속연월" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							    	<TD CLASS="TD5" NOWRAP>출력선택</TD>
				        	        <TD CLASS="TD6"><INPUT TYPE="CHECKBOX" CLASS="CHECK" NAME="prt_check1_flag" TAG="1X" VALUE="YES" ID="prt_check1_flag" ><LABEL FOR="prt_check1_flag">부표</LABEL>&nbsp;
				        	                        &nbsp;&nbsp;( T:△ , S:□ )</TD>

  								    <TD><BUTTON NAME="btnCb_autoisrt" CLASS="CLSLBTN" ONCLICK="VBScript: AutoButtonClicked('1')">자동생성</BUTTON></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>제출연월일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtsubmit_yymm NAME="txtsubmit_yymm" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="제출연월일" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
				        	        <TD CLASS="TD5">퇴직자적용일자</TD>
				        	        <TD CLASS="TD6" cols=2><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtRetireFr_dt NAME="txtRetireFr_dt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
				        							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtRetireTo_dt NAME="txtRetireTo_dt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT>
				        	        </TD>
								</TR>
					 	        <TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>		
							        <TD CLASS=TD5 NOWRAP>연말정산적용월</TD>
  								    <TD CLASS="TD6" cols=2  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtYearEnd_yymm NAME="txtYearEnd_yymm" CLASS=FPDTYYYYMM title=FPDATETIME ALT="연말정산적용월" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
								</TR>							
					        </TABLE>
				        </FIELDSET>
				    </TD>
				</TR>
				<TR>                    <!-- Condition Area-->
				    <TD <%=HEIGHT_TYPE_03%>WIDTH="100%"></TD>
				</TR>
			    <TR>	                 <!-- space Area-->
				    <TD WIDTH="100%" HEIGHT=* valign=top>
                        <TABLE <%=LR_SPACE_TYPE_60%> bgcolor=#EEEEEC>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP>
                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALGIN="LEFT">I.원천징수내역 및 납부세액</LEGEND>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
									   <TR>
									       <TD BGCOLOR=#d1e8f9 width="22%" COLSPAN="2" ROWSPAN="3">구분</TD>
									       <TD BGCOLOR=#d1e8f9 width="9%"  ROWSPAN="3">코드</TD>
									       <TD BGCOLOR=#d1e8f9 width="32%" COLSPAN="5">원천징수내역</TD>
								           <TD BGCOLOR=#d1e8f9 width="7%"  ROWSPAN="3" >6.당월조정환급세액</TD>
									       <TD BGCOLOR=#d1e8f9 width="14%" COLSPAN="2">납부세액</TD>
									  </TR>
									  <TR>
									       <TD BGCOLOR=#d1e8f9 width="11%" COLSPAN="2">소득지급(과세미달,비과세포함)</TD>
									       <TD BGCOLOR=#d1e8f9 width="21%" COLSPAN="3">징수세액</TD>
									       <TD BGCOLOR=#d1e8f9 width="7%"  ROWSPAN="2">7.소득세등(가산세포함)</TD>
									       <TD BGCOLOR=#d1e8f9 width="7%"  ROWSPAN="2">8.농어촌특별세</TD>    
									  </TR>
									  <TR>
									       <TD BGCOLOR=#d1e8f9 width="3%">1.인원</TD>
									       <TD BGCOLOR=#d1e8f9 width="7%">2.총지급액</TD>
									       <TD BGCOLOR=#d1e8f9 width="7%" >3.소득세</TD>
									       <TD BGCOLOR=#d1e8f9 width="7%" >4.농어촌특별세</TD>
									       <TD BGCOLOR=#d1e8f9 width="7%" >5.가산세</TD>
									  </TR> 
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="4%" ROWSPAN="5">근로소득</TD>
                                           <TD BGCOLOR=#d1e8f9 width="20%">간이세액</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A01</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A011 name=txt_i_A011 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A012 name=txt_i_A012 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>

                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A013 name=txt_i_A013 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A014 name=txt_i_A014 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A015 name=txt_i_A015 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%">중도퇴사</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A02</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A021 name=txt_i_A021 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A022 name=txt_i_A022 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A023 name=txt_i_A023 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A024 name=txt_i_A024 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A025 name=txt_i_A025 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>

                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%">일용근로</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A03</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A031 name=txt_i_A031 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A032 name=txt_i_A032 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A033 name=txt_i_A033 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A035 name=txt_i_A035 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%">연말정산</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A04</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A041 name=txt_i_A041 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A042 name=txt_i_A042 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A043 name=txt_i_A043 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A044 name=txt_i_A044 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A045 name=txt_i_A045 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           
                                           <TD BGCOLOR=#DDDDDD  width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD	width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD  width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%">가감계</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A10</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A101 name=txt_i_A101 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="24X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A102 name=txt_i_A102 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
  
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A103 name=txt_i_A103 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A104 name=txt_i_A104 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A105 name=txt_i_A105 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A106 name=txt_i_A106 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A107 name=txt_i_A107 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A108 name=txt_i_A108 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                 
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="24%" COLSPAN = "2">퇴직소득</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A20</TD>       
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A201 name=txt_i_A201 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A202 name=txt_i_A202 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A203 name=txt_i_A203 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A205 name=txt_i_A205 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A206 name=txt_i_A206 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A207 name=txt_i_A207 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>  

                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="4%" ROWSPAN="3">사업소득</TD>
                                           <TD BGCOLOR=#d1e8f9 width="20%">매월징수</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A25</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A251 name=txt_i_A251 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A252 name=txt_i_A252 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A253 name=txt_i_A253 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A255 name=txt_i_A255 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%">연말정산</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A26</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A261 name=txt_i_A261 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A262 name=txt_i_A262 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A263 name=txt_i_A263 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A264 name=txt_i_A264 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A265 name=txt_i_A265 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>   
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%">가감계</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A30</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A301 name=txt_i_A301 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="24X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A302 name=txt_i_A302 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
  
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A303 name=txt_i_A303 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A304 name=txt_i_A304 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A305 name=txt_i_A305 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A306 name=txt_i_A306 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A307 name=txt_i_A307 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A308 name=txt_i_A308 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>

                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">기타소득</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A40</TD>       
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A401 name=txt_i_A401 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A402 name=txt_i_A402 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A403 name=txt_i_A403 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A405 name=txt_i_A405 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A406 name=txt_i_A406 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A407 name=txt_i_A407 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">연금소득</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A45</TD>       
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A451 name=txt_i_A451 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A452 name=txt_i_A452 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A453 name=txt_i_A453 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A455 name=txt_i_A455 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A456 name=txt_i_A456 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A457 name=txt_i_A457 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>  
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">이자소득</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A50</TD>       
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_i_A501 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_i_A502 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A503 name=txt_i_A503 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A504 name=txt_i_A504 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A505 name=txt_i_A505 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A506 name=txt_i_A506 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A507 name=txt_i_A507 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A508 name=txt_i_A508 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>  
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">배당소득</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A60</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A601 name=txt_i_A601 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A602 name=txt_i_A602 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A603 name=txt_i_A603 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A604 name=txt_i_A604 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A605 name=txt_i_A605 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A606 name=txt_i_A606 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A607 name=txt_i_A607 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A608 name=txt_i_A608 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>                                                                                                                    
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">저축해지추징세액</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A69</TD>       
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A691 name=txt_i_A691 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A693 name=txt_i_A693 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A694 name=txt_i_A694 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A695 name=txt_i_A695 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A696 name=txt_i_A696 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A697 name=txt_i_A697 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A698 name=txt_i_A698 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>  
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">법인원천</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A80</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A801 name=txt_i_A801 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="21X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A802 name=txt_i_A802 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A803 name=txt_i_A803 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A805 name=txt_i_A805 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A806 name=txt_i_A806 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A807 name=txt_i_A807 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                      </TR>  
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">수정신고(세액)</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A90</TD>
                                           <TD BGCOLOR=#DDDDDD   width="3%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           <TD BGCOLOR=#DDDDDD   width="7%"><TABLE BGCOLOR=#EEEEEC width=100% height=100%><TR><TD BGCOLOR=#EEEEEC></TD></TR></TABLE></TD>
                                           
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A903 name=txt_i_A903 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A904 name=txt_i_A904 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A905 name=txt_i_A905 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A906 name=txt_i_A906 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A907 name=txt_i_A907 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A908 name=txt_i_A908 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="20%" COLSPAN = "2">총합계</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A99</TD>
                                           <TD width="3%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A991 name=txt_i_A991 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 48px" title=FPDOUBLESINGLE tag="24X6Z" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A992 name=txt_i_A992 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A993 name=txt_i_A993 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A994 name=txt_i_A994 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A995 name=txt_i_A995 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="6%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A996 name=txt_i_A996 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 84px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A997 name=txt_i_A997 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="7%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txt_i_A998 name=txt_i_A998 style="HEIGHT: 17px; LEFT: 0px; TOP: 0px; WIDTH: 94px" title=FPDOUBLESINGLE tag="24X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
                                   </FIELDSET>

                                   <FIELDSET CLASS="CLSFLD"><LEGEND ALGIN="LEFT">II.환급세액조정</LEGEND>
                                   <table <%=LR_SPACE_TYPE_20%> border="1" width="100%">
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="30%" COLSPAN="3">전월미환급세액의계산</TD>
                                           <TD BGCOLOR=#d1e8f9 width="30%" COLSPAN="3">D.당월발생환급세액</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" COLSPAN="2" ROWSPAN="2">E.조정대상환급세액(C+D)</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" COLSPAN="2" ROWSPAN="2">F.당월조정환급세액</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%" COLSPAN="2" ROWSPAN="2">G.차월이월(E-F)</TD>
                                      </TR>
                                      <TR>
                                           <TD BGCOLOR=#d1e8f9 width="9%">A.전월미환급세액</TD>
                                           <TD BGCOLOR=#d1e8f9 width="11%">B.기환급신청한세액</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%">C.차감잔액(A-B)</TD>
                                           <TD BGCOLOR=#d1e8f9 width="9%">1.일반환급(C+D)</TD>
                                           <TD BGCOLOR=#d1e8f9 width="11%">2.신탁재산(금융기관)</TD>
                                           <TD BGCOLOR=#d1e8f9 width="10%">3.(기타)</TD>
                                      </TR>                                      
                                      <TR>
                                           <TD width="9%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A001 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="11%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A002 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 110px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A003 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="9%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A004 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="11%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A005 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 120px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A006 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 80px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" COLSPAN="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A007 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 110px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" COLSPAN="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A008 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 110px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                           <TD width="10%" COLSPAN="2"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=hfa050t_paria_cnt name=txt_ii_A009 style="HEIGHT: 16px; LEFT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE tag="21X2" ></OBJECT>');</SCRIPT></TD>
                                      </TR>
                                   </TABLE>
                                   </FIELDSET>
                                </TD>
                              </TR>
                        </TABLE>
                     </TD>
                 </TR>
<!-- Space Area -->
	
<!-- Button, Batch, Print, Jump Area -->
            </TABLE>
        </TD>
    </TR>
	<TR >
	    <TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP"  SRC = "../../blank.htm"  WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
	</TR>
<INPUT TYPE=HIDDEN NAME="txtMode"        Tag="21">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   Tag="21">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  Tag="21">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" Tag="21">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     Tag="21">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    Tag="21">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>
