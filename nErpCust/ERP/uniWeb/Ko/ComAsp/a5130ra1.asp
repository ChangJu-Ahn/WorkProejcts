<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5103RA1
'*  4. Program Name         : 
'*  5. Program Desc         : Ado query Sample with DBAgent(Multi + Multi)
'*  6. Component List       :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2002/02/23
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : Park ShimSeo
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :  2002/11/25 : ASP Standard for Include improvement
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../inc/incSvrCcm.inc"  -->
<!-- #Include file="../inc/incSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAMain.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAEvent.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliPAOperation.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliVariables.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE ="JavaScript"SRC = "../inc/incImage.js">					</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../inc/incEB.vbs">						</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit 
	
Dim arrParent
Dim arrParam	
'------ Set Parameters from Parent ASP -----------------------------------------------------------------------
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)
	
top.document.title = PopupParent.gActivePRAspName	
	
Const C_MASTER = 1
Const C_DETAIL = 2
	

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID        = "a5130rb1_ko441.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a5130rb2.asp"                         '☆: Biz logic spread sheet for #2

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_MaxKey            = 5                                    '☆☆☆☆: Max key value

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgDocCur                          
Dim lgPageNo_A
Dim lgPageNo_B
Dim lgSortKey_A
Dim lgSortKey_B
Dim lgIsOpenPop


'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgBlnFlgChgValue	= False 
    lgIntFlgMode		= PopupParent.OPMD_CMODE

    lgPageNo_A			= ""
    lgSortKey_A			= 1

    lgPageNo_B			= ""
    lgSortKey_B			= 1

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
	frm1.txttempGlNo.value  = arrParam(0)
	frm1.txtRefNo.value = arrParam(1)
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

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
	<!-- #Include file="./LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","RA") %>
End Sub


'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function


Sub InitSpreadSheet()
    
	Call SetZAdoSpreadSheet("A5130RA1", "S", "A", "V20030210", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetZAdoSpreadSheet("A5130RA1_DTL", "S", "B", "V20030210", PopupParent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X")
	Call SetSpreadLock ("A")
	Call SetSpreadLock ("B")
    
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock( iOpt )
    If iOpt = "A" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData 
          ggoSpread.SpreadLockWithOddEvenRowColor()
          .vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2 
			ggoSpread.SpreadLockWithOddEvenRowColor()
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub



'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
	frm1.bttnGlRefView.disabled = True
    '--------- Developer Coding Part (End  ) ----------------------------------------------------------

	If GetDucCur() Then
		Call CurFormatNumericOCX()
	End If

	Call FncQuery

End Sub


'========================================================================================================
'   Event Name : GetDucCur()
'   Event Desc :
'========================================================================================================
Function GetDucCur()
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    Dim strBizAreaCd, strBizAreaNm
    Dim strSelect
    Dim strFrom
    Dim strWhere
    Dim arrTemp
    
    GetDucCur = False
    strSelect	= "isnull(doc_cur,'')"
    strFrom		= "a_temp_gl_item"
    strWhere	= "temp_gl_no='" & frm1.txttempGlNo.value & "'" 
    
    If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrTemp		= split(lgF0, Chr(11))
		lgDocCur	= arrTemp(0) 		
		if Trim(lgDocCur) = "" Then
			GetDucCur = False
		Else
			GetDucCur = True
		End If					
	End If
End Function

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    Dim IntRetCD

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
     If Trim(frm1.txttempGlNo.value) = "" And Trim(frm1.txtRefNo.value) = "" Then
		Call DisplayMsgBox("114100", "X", "X", "X")
		Call CancelClick()
		Exit Function
    End If
	
    '-----------------------
    'Query function call area
    '-----------------------
	frm1.vspdData.MaxRows = 0                                                      '☜: Protect system from crashing
    Call DbQuery(C_MASTER)															'☜: Query db data

    FncQuery = True	
End Function


'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    Dim iColumnLimit2
    
    If gMouseClickStatus = "SPCRP" Then
       iColumnLimit  =3
       
       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData.ScrollBars = PopupParent.SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow
    
       Frm1.vspdData.Action = 0    
    
       Frm1.vspdData.ScrollBars = PopupParent.SS_SCROLLBAR_BOTH
    End If   
	
	'----------------------------------------
	' Spread가 두개일 경우 2번째 Spread
	'----------------------------------------
	
	
    If gMouseClickStatus = "SP2CRP" Then
		iColumnLimit2 = 4
       
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit2 Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit2 , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData2.ScrollBars = PopupParent.SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData2
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow
    
       Frm1.vspdData2.Action = 0    
    
       Frm1.vspdData2.ScrollBars = PopupParent.SS_SCROLLBAR_BOTH
    End If   
    
End Function

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo)
	Dim intRetCd

	StrEbrFile = "a5118ma1"
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txttempGlDt.value, PopupParent.gDateFormat, PopupParent.gServerDateType)	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txttempGlDt.value, PopupParent.gDateFormat, PopupParent.gServerDateType)

' 회계전표의 key는 GL_NO이기 때문에 GL_NO만 넘긴다.	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	VarTempGlNoFr = Trim(frm1.txttempGlNo.value)
	VarTempGlNoTo = Trim(frm1.txttempGlNo.value)
	
End Sub
''======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo, varGlPutType
    Dim StrEbrFile
    Dim intRetCd
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	varGlPutType ="%"

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo)
	
'    On Error Resume Next                                                    '☜: Protect system from crashing
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	
	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|GlPutType|" & varGlPutType	
	StrUrl = StrUrl & "|OrgChangeId|" & PopupParent.gChangeOrgId
	
	StrUrl = StrUrl & "|gUsrId|" & PopupParent.gUsrId		'>>air	
	StrUrl = StrUrl & "|LoginDeptNm|" & PopupParent.gDepart	'>>air			
	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo, varGlPutType
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	varGlPutType ="%"

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo)
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|GlPutType|" & varGlPutType	
	StrUrl = StrUrl & "|OrgChangeId|" & PopupParent.gChangeOrgId
	
	StrUrl = StrUrl & "|gUsrId|" & PopupParent.gUsrId		'>>air	
	StrUrl = StrUrl & "|LoginDeptNm|" & PopupParent.gDepart	'>>air		
	
	Call FncEBRPreview(ObjName,StrUrl)
	
End Function

Function FncBtnGlRefPopUp()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	

	If lgIsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra2")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", popupparent.VB_INFORMATION, "a5130ra2", "X")
		lgIsOpenPop = False
		Exit Function
	End If	

	arrParam(0) = Trim(frm1.hHqBrchNo.value)	'결의전표번호 
	'arrParam(1) = ""			'Reference번호 
	

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.popupparent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect) 
	Dim strVal

    DbQuery = False
    
    Err.Clear 
    
    Select Case pDirect
		Case  C_MASTER           			
			Call LayerShowHide(1)
			    
			With frm1
				strVal = BIZ_PGM_ID & "?txttempglno=" & Trim(.txttempglno.value)
				strVal = strVal & "&txtRefNo=" & Trim(.txtRefNo.value)
				strVal = strVal & "&txttempglno_Alt=" & Trim(.txttempglno.Alt)
				strVal = strVal & "&txtRefNo_Alt=" & Trim(.txtRefNo.Alt)
			    strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag								
				strVal = strVal & "&lgPageNo="       & lgPageNo_A                          '☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
				strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

			    Call RunMyBizASP(MyBizASP, strVal)	
			End With
    
		Case C_Detail
			frm1.vspdData2.MaxRows = 0
			Call LayerShowHide(1)
			    
			With frm1
'---
			
				strVal = BIZ_PGM_ID1 & "?txtTempGlNo=" & Trim(GetKeyPosVal("A", 1))
				strVal = strVal & "&txtSeq=" & Trim(GetKeyPosVal("A", 2))
'---
			    strVal = strVal & "&lgPageNo="       & lgPageNo_B                          '☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")         
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
			    
			    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
			End With
	End Select 
    DbQuery = True   

End Function


'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk( iOpt)											 '☆: 조회 성공후 실행로직 
	
	
	Dim lngRows
	Dim strTableid
	Dim strColid
	Dim strColNm
	Dim strMajorCd
	Dim strNmwhere
	Dim arrVal
	
	Const C_Tableid		= 8
	Const C_Colid		= 9
	Const C_ColNm		= 10
	Const C_MajorCd		= 14
	Const C_CtrlVal		= 4
	Const C_CtrlValNm	= 6


	lgIntFlgMode     = PopupParent.OPMD_UMODE
	
	'이체번호가 있을 경우 관련 전표 버튼을 enable 한다.
    if trim(frm1.hHqBrchNo.value) <> "" then
		frm1.bttnGlRefView.disabled	=	false	
	end if
	

	If iOpt = 1 Then
       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If
    
    For lngRows = 1 To frm1.vspdData2.Maxrows
		frm1.vspddata2.row = lngRows	
		frm1.vspddata2.col = C_Tableid 
		IF Trim(frm1.vspddata2.text) <> "" Then

			frm1.vspddata2.col = C_Tableid
			strTableid = frm1.vspddata2.text
			frm1.vspddata2.col = C_Colid
			strColid = frm1.vspddata2.text
			frm1.vspddata2.col = C_ColNm
			strColNm = frm1.vspddata2.text	
			frm1.vspddata2.col = C_MajorCd
			strMajorCd = frm1.vspddata2.text

			frm1.vspddata2.col = C_CtrlVal

			strNmwhere = strColid & " =  '" & frm1.vspddata2.text & "' " 

			IF Trim(strMajorCd) <> "" Then
				strNmwhere = strNmwhere & " AND MAJOR_CD = '" & strMajorCd & "' "
			End IF
			IF CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
				frm1.vspddata2.col = C_CtrlValNm
				arrVal = Split(lgF0, Chr(11))  
				If Ubound(arrVal, 1) <> - 1 Then 
					frm1.vspddata2.text = arrVal(0)
				End If
			End IF
		End IF
	Next

	Call ggoOper.LockField(Document, "Q")
End Function

'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenSortPopup()

	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("./ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click( Col,  Row)
    Dim ii

	gMouseClickStatus = "SPC"

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 1
        End If
        Exit Sub
    End If

	If Row < 1 Then Exit Sub
		
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)	
    Call DbQuery(C_DETAIL)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    ggoSpread.Source = frm1.vspdData
		    If lgSortKey_A = 1 Then
		        ggoSpread.SSSort, lgSortKey_A
		        lgSortKey_A = 2
		    Else
		        ggoSpread.SSSort, lgSortKey_A
		        lgSortKey_A = 1
		    End If    
		    Exit Sub
		End If

		If Col < 1 Then Exit Sub
		
		Call SetSpreadColumnValue("A",frm1.vspdData,Col,NewRow)
		Call DbQuery(C_DETAIL)
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click( Col,  Row)
    Dim ii
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If

	gMouseClickStatus = "SP2C"

End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           If DbQuery(C_MASTER) = False Then
              Exit Sub
           End if
		End If
   End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           If DbQuery(C_DETAIL) = False Then
              Exit Sub
          End if
		End If
   End if
    
End Sub

'========================================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'=========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 then
		frm1.fpdtFromEnterDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpdtFromEnterDt.Focus
	End if
End Sub
'========================================================================================================
'   Event Name : txtPoToDt
'   Event Desc :
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 then
		frm1.fpdtToEnterDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpdtToEnterDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : fpdtFromEnterDt_KeyDown(keycode, shift)
'   Event Desc : 
'=======================================================================================================
Sub fpdtFromEnterDt_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'=======================================================================================================
'   Event Name : fpdtToEnterDt_KeyDown(keycode, shift)
'   Event Desc : 
'=======================================================================================================
Sub fpdtToEnterDt_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'차변금액 

		ggoOper.FormatFieldByObjectOfCur .txtDrAmt, lgDocCur, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
		'대변금액 
		ggoOper.FormatFieldByObjectOfCur .txtCrAmt, lgDocCur, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
	End With

End Sub


</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>


<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
				<TR>
					<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txttempGlNo" MAXLENGTH="18" SIZE=20  ALT ="결의전표번호" tag="14XXXU"></TD>
					<TD CLASS=TD5 NOWRAP>참조번호</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" MAXLENGTH="30" SIZE=30  ALT ="참조번호" tag="14XXXU"></TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP>결의전표일자</TD>
					<TD CLASS="TD6" NOWRAP><INPUT NAME="txttempGlDt" ALT="결의전표일자" SIZE = "10" MAXLENGTH="10" STYLE="TEXT-ALIGN: Center" tag="24X1"></TD>
					<TD CLASS="TD5" NOWRAP>부서</TD>
					<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" Size= "10" MAXLENGTH="10"  tag="24XXXU" >&nbsp;<INPUT NAME="txtDeptNm" ALT="부서명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>차대합계</TD>
					<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5130ra1_I992785608_txtDrAmt.js'></script>&nbsp;
										 <script language =javascript src='./js/a5130ra1_I328279000_txtCrAmt.js'></script></TD>
					<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
					<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5130ra1_I205996823_txtDrLocAmt.js'></script>&nbsp;
										 <script language =javascript src='./js/a5130ra1_I338999360_txtCrLocAmt.js'></script></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>적요</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlDesc" MAXLENGTH="30" SIZE=32 ALT ="적요" tag="14XXXU"></TD>
					<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGlInputType" ALT="전표입력경로" Size= "10" MAXLENGTH="10" tag="24XXXU" >&nbsp;<INPUT NAME="txtGlInputTypeNm" ALT="전표입력경로명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="24X"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>승인여부</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConfFgNm" MAXLENGTH="10" SIZE=10 ALT ="승인여부" tag="14XXXU"></TD>
					<TD CLASS=TD5 NOWRAP>승인자</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConfirmEmp" MAXLENGTH="30" SIZE=10 ALT ="승인자" tag="14XXXU"></TD>
				</TR>

				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<script language =javascript src='./js/a5130ra1_vspdData_vspdData.js'></script>
					</TD>
				</TR>
				<TR HEIGHT="40%">
					<TD>
						<script language =javascript src='./js/a5130ra1_vspdData2_vspdData2.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				<TD >&nbsp;&nbsp;<BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
				                 <BUTTON NAME="bttnPrint"	 CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()"   Flag=1>인쇄</BUTTON>&nbsp;
				                 <BUTTON NAME="bttnGlRefView"	 CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnGlRefPopUp()"   Flag=0>관련전표조회</BUTTON></TD>
				<TD ALIGN=RIGHT> <IMG SRC="../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Query.gif',1)" >	 </IMG>&nbsp;
								 <IMG SRC="../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/zpConfig.gif',1)" ></IMG>&nbsp;
                                 <IMG SRC="../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/Cancel.gif',1)">	 </IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>		
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hHqBrchNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>