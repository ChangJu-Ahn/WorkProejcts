<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : Billing
'*  3. Program ID           : S5111QA4
'*  4. Program Name         : 영업조직별 월 매출현황(T)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2002/10/16
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">                    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Common.vbs">                 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Event.vbs">                  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Variables.vbs">              </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Operation.vbs">              </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/AdoQuery.vbs">               </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgent.vbs">          </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<Script Language="vbscript"	  src="../../inc/incUni2KTV.vbs"></Script>
<Script Language="VBScript">
Option Explicit                                                  
	
' External ASP File
'========================================
Const BIZ_PGM_ID        = "S5111QB4.asp"                         

' Constant variables 
'========================================
Const C_MaxKey            = 11

' Tree view 관련 추가 
Const  C_Root = "Root"
Const  C_ORG = "ORG"
Const  C_GRP = "GRP"
Const  C_ORG_SUFFIX = "O"		' This must be one character
CONST  C_GRP_SUFFIX = "G"

Const  C_ROOT_DESC = "UNIERP"
Const  C_ROOT_KEY = "$"
Const  C_ROOT_KEY_STR = "RT_"
Const  C_UNDERSCORE = "_"

Const C_IMG_Root = "../../image/unierp.gif"
Const C_IMG_ORG = "../../image/Orglvl_2.gif"
Const C_IMG_Open = "../../image/Group_op.gif"
Const C_IMG_GRP = "../../image/HumanC.gif"
Const C_IMG_None = "../../image/c_none.gif"
Const C_IMG_Const = "../../image/c_const.gif"

Const C_PopSalesOrg	= 1

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgIsOpenPop        
Dim IsOpenPop                                       '☜: Popup status                           

'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim lgSelectList_A                                          '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgSelectListDT_A                                        '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgPopUpR_A                                              '☜: Orderby,Groupby default 값            

Dim lgSortFieldNm_A                                         '☜: Orderby popup용 데이타(필드설명)      
Dim lgSortFieldCD_A                                         '☜: Orderby popup용 데이타(필드코드)      

Dim lgPageNo_A                                                                        
Dim lgSortKey_A                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim lgSelectList_B                                          '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgSelectListDT_B                                        '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgPopUpR_B                                              '☜: Orderby,Groupby default 값            

Dim lgSortFieldNm_B                                         '☜: Orderby popup용 데이타(필드설명)      
Dim lgSortFieldCD_B                                         '☜: Orderby popup용 데이타(필드코드)      

Dim lgPageNo_B                                                                        
Dim lgSortKey_B                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet temp---------------------------------------------------------------------------   
Dim lgTypeCD_T                                              '☜: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD_T                                             '☜: 필드 코드값                           
Dim lgFieldNM_T                                             '☜: 필드 설명값                           
Dim lgFieldLen_T                                            '☜: 필드 폭(Spreadsheet관련)              
Dim lgFieldType_T                                           '☜: 필드 설명값                           
Dim lgDefaultT_T                                            '☜: 필드 기본값                           
Dim lgNextSeq_T                                             '☜: 필드 Pair값                           
Dim lgKeyTag_T                                              '☜: Key 정보                              

Dim lgSelectList_T                                          '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgSelectListDT_T                                        '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgPopUpR_T                                              '☜: Orderby,Groupby default 값            

Dim lgSortFieldNm_T                                         '☜: Orderby popup용 데이타(필드설명)      
Dim lgSortFieldCD_T                                         '☜: Orderby popup용 데이타(필드코드)      

Dim lgKeyPos                                                '☜: Key위치 
' 1 : Year 
' 2 : Month
' 3 : Sales Org.
' 4 : Sales Org. Name
' 5 : Billing Amt.
' 6 : VAT Amt.
' 7 : Billing Amt + VAT Amt.
' 8 : Sales Org + Suffix
' 9 : Parent Sales Org. + Suffix
                               
Dim lgKeyPosVal                                             '☜: Key위치 Value                         

Dim lgBlnOpenedFlag
Dim	lgBlnSalesOrgChg
Dim lgBlnFlgConChg
Dim	lgStrPrevNodeKey
Dim lgStrRootKey
Dim lgStrRootDesc

<%                                                  
	BaseDate = GetSvrDate                                                                  'Get DB Server Date
	EndDate = UNIConvDateAtoB(BaseDate, gServerDateFormat, gDateFormatYYYYMM)
%>

'========================================
Sub InitVariables()

    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgIntFlgMode     = OPMD_CMODE                          

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B   = ""										'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1

	lgStrPrevKey = ""										'initializes Previous Key
	lgStrPrevNodeKey = ""
	lgStrRootKey = ""
	lgStrRootDesc = ""
	lgBlnSalesOrgChg = False								' 영업조직변경여부 
    lgBlnFlgConChg	 = False
End Sub

'========================================
Sub SetDefaultVal()

	With frm1
		.txtFromDt.Text = "<%=EndDate%>"
		.txtToDt.Text = "<%=EndDate%>"	
		.cboQueryData.value = "B"
		.cboSalesOrgLvl.value = 1
		.txtFromDt.Focus
	End With
End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "Q", "S") %>

End Sub

'========================================
Sub InitTree()
    With frm1.uniTree1
		.HideSelection = false
		.SetAddImageCount = 6
		.Indentation = "200"	' 줄 간격 
						' 파일위치,	키명, 위치 
		.AddImage C_IMG_Root,		C_Root,		0
		.AddImage C_IMG_ORG,		C_ORG,		0
		.AddImage C_IMG_Open,		C_Open,		0
		.AddImage C_IMG_GRP,		C_GRP,		0
		.AddImage C_IMG_None,		C_None,		0
		.AddImage C_IMG_Const,		C_Const,	0
	
		.PathSeparator = gColSep
		
		.OLEDragMode = 0														'⊙: Drag & Drop 을 가능하게 할 것인가 정의 
		.OLEDropMode = 0	
	
	End With
End Sub		

'========================================
Sub InitSpreadSheet(ByVal pOpt)
    Dim iMaxColumn
    Dim lgMaxFieldCount
    
    lgSelectList_T   = ""
    lgSelectListDT_T = ""
    iMaxColumn       = 0 
    
    lgMaxFieldCount = UBound(lgFieldNM_T)

    If pOpt = "1" Then                                   ' 초기화 Spreadsheet #1 
       ggoSpread.Source = Frm1.vspdData
       With frm1.vspdData
          .MaxCols = 0
          .MaxCols = lgMaxFieldCount
          .MaxRows = 0
'          .OperationMode = 2
          .ReDraw = false
       End With 

       With frm1.vspdData2
          .MaxRows = 0
       End With 

    Else                                                ' 초기화 Spreadsheet #2 
       ggoSpread.Source = Frm1.vspdData2
       With frm1.vspdData2
          .MaxCols = 0
          .MaxCols = lgMaxFieldCount
          .MaxRows = 0
'          .OperationMode = 3
          .ReDraw = false
       End With 
    End If   
    
    ggoSpread.Spreadinit

    Call CopyToTmpBuffer(lgTypeCD_T,lgFieldCD_T,lgFieldNM_T,lgFieldLen_T,lgFieldType_T,lgDefaultT_T,lgNextSeq_T,lgKeyTag_T)
        
    If pOpt = "1" Then                                   ' 초기화 Spreadsheet #1 
       iMaxColumn = InitSpreadSheetFieldOfZADO(frm1.vspdData ,lgPopUpR_T,lgSelectList_T,lgSelectListDT_T,lgKeyPos,C_MaxKey,C_MaxSelList)
    Else
       iMaxColumn = InitSpreadSheetFieldOfZADO(frm1.vspdData2,lgPopUpR_T,lgSelectList_T,lgSelectListDT_T,lgKeyPos,C_MaxKey,C_MaxSelList)
    End If   

    If pOpt = "1" Then
	   ggoSpread.SSSetSplit(1)											'frozen 기능 추가 
       ggoSpread.Source = Frm1.vspdData
       With frm1.vspdData
         .MaxCols = iMaxColumn
         .ReDraw = true
       End With 
    Else
	   ggoSpread.SSSetSplit(2)											'frozen 기능 추가 
       ggoSpread.Source = Frm1.vspdData2
       With frm1.vspdData2
         .MaxCols = iMaxColumn
         .ReDraw = true
       End With 
    End If   

    Call SetSpreadLock (pOpt)
    Call CopyPopupInfTAB(pOpt)
    
End Sub

'========================================
Sub SetSpreadLock(ByVal pOpt )
    If pOpt = "1" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData 
          ggoSpread.SpreadLock 1 , -1
          .vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2 
            ggoSpread.SpreadLock 1, -1
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub

'========================================
Sub SetPopUpInitialInf(ByVal pOpt)

    ReDim lgPopUpR_T(C_MaxSelList - 1,1)
    
    Call MakePopData(lgDefaultT_T,lgFieldNM_T,lgFieldCD_T,lgPopUpR_T,lgSortFieldNm_T,lgSortFieldCD_T,C_MaxSelList)
    
    If pOpt = "1" Then          
       lgSortFieldCD_A = lgSortFieldCD_T                      '배열화 
       lgSortFieldNM_A = lgSortFieldNm_T
    Else
       lgSortFieldCD_B = lgSortFieldCD_T
       lgSortFieldNM_B = lgSortFieldNm_T       
    End If       
    
End Sub 

'========================================
Sub CopyPopupInfABT(ByVal pOpt)

    Call CopyTBL(pOpt)    

    If pOpt = "1" Then
       lgPopUpR_T      = lgPopUpR_A
       lgSortFieldCD_T = lgSortFieldCD_A
       lgSortFieldNM_T = lgSortFieldNM_A       
    Else
       lgPopUpR_T      = lgPopUpR_B
       lgSortFieldCD_T = lgSortFieldCD_B
       lgSortFieldNM_T = lgSortFieldNM_B
    End If       
End Sub

'========================================
Sub CopyPopupInfTAB(ByVal pOpt)

    If pOpt = "1" Then
       lgPopUpR_A       = lgPopUpR_T
       lgSelectList_A   = lgSelectList_T  
       lgSelectListDT_A = lgSelectListDT_T
    Else
       lgPopUpR_B       = lgPopUpR_T
       lgSelectList_B   = lgSelectList_T  
       lgSelectListDT_B = lgSelectListDT_T
    End If       
End Sub

'========================================
Sub CopyTBL(ByVal pOpt)


   Select Case pOpt
      Case "1"
              lgTypeCD_T    = gTypeCD
              lgFieldCD_T   = gFieldCD
              lgFieldNM_T   = gFieldNM
              lgFieldLen_T  = gFieldLen
              lgFieldType_T = gFieldType
              lgDefaultT_T  = gDefaultT
              lgNextSeq_T   = gNextSeq
              lgKeyTag_T    = gKeyTag
      Case "2"
              lgTypeCD_T    = gTypeCD1
              lgFieldCD_T   = gFieldCD1
              lgFieldNM_T   = gFieldNM1
              lgFieldLen_T  = gFieldLen1
              lgFieldType_T = gFieldType1
              lgDefaultT_T  = gDefaultT1
              lgNextSeq_T   = gNextSeq1
              lgKeyTag_T    = gKeyTag1
    End Select              
End Sub

'========================================
Sub Form_Load()
	on Error Resume Next
	
	Call LoadInfTB19029														
    Call GetAdoFieldInf("S5111QA4","S","A")                                    ' S for Sort , A for SpreadSheet No('A','B',....
    Call GetAdoFieldInf("S5111QA4","S","B")                                    ' S for Sort , A for SpreadSheet No('A','B',....

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFromDt, gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtToDt, gDateFormat, 2)
    
    Call ggoOper.LockField(Document, "N")                                      ' ⊙: Lock  Suitable  Field
    
    ReDim lgPopUpR_A(C_MaxSelList - 1,1)
    ReDim lgPopUpR_B(C_MaxSelList - 1,1)
    ReDim lgKeyPos(C_MaxKey)
    ReDim lgKeyPosVal(C_MaxKey)

    Call InitComboBox()
	Call InitVariables														
	Call SetDefaultVal	
    Call CopyTBL("1")                                                       '⊙: Initializes Spread Sheet #1
    Call SetPopUpInitialInf("1")
	Call InitSpreadSheet("1")
	
    Call CopyTBL("2")                                                       '⊙: Initializes Spread Sheet #2
    Call SetPopUpInitialInf("2")
	Call InitSpreadSheet("2")
	Call InitTree()
	frm1.vspdDataH.MaxCols = frm1.vspdData.MaxCols
	frm1.vspdDataH2.MaxCols = frm1.vspdData.MaxCols

	lgBlnOpenedFlag = True

    Set gActiveElement = document.activeElement 
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'========================================
Function FncQuery() 

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncQuery = False                                                              

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								
       Exit Function
    End If

	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
   
	' 조회조건 유효값 check
	If 	lgBlnFlgConChg Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If

    Call ggoOper.ClearField(Document, "2")								          
	frm1.uniTree1.Nodes.Clear
    
    Call InitVariables                                                            
    

	If DbQuery = False Then   
       Exit Function           
    End If     							


    If Err.number = 0 Then
       FncQuery = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncPrint()

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncPrint = False                                                              

    If Err.number = 0 Then
       FncPrint = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExcel() 

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncExcel = False                                                             

	Call Parent.FncExport(C_MULTI)

    If Err.number = 0 Then
       FncExcel = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncFind() 

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncFind = False                                                               

	Call Parent.FncFind(C_MULTI, True)

    If Err.number = 0 Then
       FncFind = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    Dim iColumnLimit2
    
    If gMouseClickStatus = "SPCRP" Then
       iColumnLimit  = frm1.vspdData.MaxCols
       
       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow
    
       Frm1.vspdData.Action = 0    
    
       Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
    End If   
	
	'----------------------------------------
	' Spread가 두개일 경우 2번째 Spread
	'----------------------------------------
	
	
    If gMouseClickStatus = "SP2CRP" Then
		iColumnLimit2 = frm1.vspdData2.MaxCols
       
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit2 Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit2 , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData2.ScrollBars = SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData2
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow
    
       Frm1.vspdData2.Action = 0    
    
       Frm1.vspdData2.ScrollBars = SS_SCROLLBAR_BOTH
    End If   
    
End Function

'========================================
Function FncExit()

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncExit = True                                                             

End Function

'========================================
Function DbQuery() 
	Dim strVal
	
    On Error Resume Next                                                          
    Err.Clear                                                                     

    DbQuery = False                                                              

    Call DisableToolBar(TBC_QUERY)
    Call LayerShowHide(1)
    Call CopyPopupInfABT("1")


    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & UID_M0001
		strVal = strVal & "&txtFromDt=" & UNIGetFirstDay(.txtFromDt.text, gDateFormatYYYYMM)
		strVal = strVal & "&txtToDt=" & UNIGetLastDay(.txtToDt.text, gDateFormatYYYYMM)
		strVal = strVal & "&txtQueryData=" & Trim(.cboQueryData.value)
		strVal = strVal & "&txtSalesOrgLvl=" & Trim(.cboSalesOrgLvl.value)
		strVal = strVal & "&txtSalesOrg=" & Trim(.txtSalesOrg.value)
		If .rdoBLFlgA.checked Then
			strVal = strVal & "&txtBlFlag=%"
		ElseIf .rdoBLFlgN.checked Then
			strVal = strVal & "&txtBlFlag=N"
		Else
			strVal = strVal & "&txtBlFlag=Y"
		End If
		
		If Len(Trim(.txtSalesOrg.value)) Then
			lgStrRootKey = Left(.txtSalesOrg.value & "   ", 4) & C_ORG_SUFFIX
			lgStrRootDesc = "[" & Left(.txtSalesOrg.value & "   ", 4) & "]" & .txtSalesOrgNm.value
		Else
			lgStrRootKey = C_ROOT_KEY
			lgStrRootDesc = C_ROOT_DESC
		End If
		strVal = strVal & "&txtRootKey=" & lgStrRootKey
		strVal = strVal & "&txtOrgSuffix=" & C_ORG_SUFFIX
		strVal = strVal & "&txtGrpSuffix=" & C_GRP_SUFFIX

        strVal = strVal      & "&lgPageNo="          & lgPageNo_A                          
        strVal = strVal      & "&lgSelectListDT="    & lgSelectListDT_A
        strVal = strVal      & "&lgTailList="        & MakeSQLGroupOrderByList(UBound(lgFieldNM_T),lgPopUpR_T,lgFieldCD_T,lgNextSeq_T,lgTypeCD_T(0),C_MaxSelList)
        strVal = strVal      & "&lgSelectList="      & EnCoding(lgSelectList_A)
		
	End With    
    
	Call RunMyBizASP(MyBizASP, strVal)

    If Err.number = 0 Then
       DbQuery = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()
	On Error Resume Next
	
	With frm1
		If .vspdDataH.MaxRows > 0 Then
			Call DisplayNodes()
			Call CopyVspdDataHToVspdDataH2()
			Call SortvspdDataH()
			Call SortvspdDataH2()

			Call uniTree1_NodeClick(.uniTree1.Nodes(lgStrRootKey))

			.uniTree1.Focus
			
			.vspdData.SelModeSelected = True
			If .vspdData.MaxRows > 0 Then			
				.vspdData.Row = 1
			End If
			lgIntFlgMode = OPMD_UMODE
		Else
			.txtFromDt.focus
		End If
	End With
End Function

'====================================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnSalesOrgChg Then
		iStrCode = Trim(frm1.txtSalesOrg.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", frm1.cboSalesOrgLvl.value, "" & FilterVar("SO", "''", "S") & "", C_PopSalesOrg) Then
				frm1.txtSalesOrg.value = ""
				frm1.txtSalesOrgNm.value = ""
'				Call DisplayMsgBox("970000", "X", frm1.txtSalesOrg.alt, "X")
				frm1.txtSalesOrg.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesOrgNm.value = ""
		End If
		lgBlnSalesOrgChg	= False
	End If

End Function

'====================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False

	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'====================================================
Sub SortvspdDataH()
	
	With frm1.vspdDataH
		.Row = 1
		.Col = 1
		.Row2 = .MaxRows
		.Col2 = .MaxCols
		' Set sort definition for key 1
		.SortBy = 0 'SS_SORT_BY_ROW
		.SortKey(1) = lgKeyPos(9)
		.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
		' Set sort definition for key 2
		.SortKey(2) = lgKeyPos(11)
		.SortKeyOrder(2) = 1 'SS_SORT_ORDER_DESCENDING
		.SortKey(3) = lgKeyPos(3)
		.SortKeyOrder(3) = 1 'SS_SORT_ORDER_ASCENDING
		.Action = 25 'SS_ACTION_SORT 
	End With
End Sub

'====================================================
Sub SortvspdDataH2()
	Dim iArrSortKeys, iArrSortKeyOrder
	
	With frm1.vspdDataH2
		.Row = 1
		.Col = 1
		.Row2 = .MaxRows
		.Col2 = .MaxCols
		' Set sort definition for key 1
		.SortBy = 0 'SS_SORT_BY_ROW
		.SortKey(1) = lgKeyPos(8)
		.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
		' Set sort definition for key 2
		.SortKey(2) = lgKeyPos(11)
		.SortKeyOrder(2) = 1 'SS_SORT_ORDER_DESCENDING
		.Action = 25 'SS_ACTION_SORT 
	End With
End Sub

'==========================================
'   Event Name : DisplayNodes
'   Event Desc : 
'==========================================

Sub DisplayNodes()
	Dim iObjDummyNode
	Dim iStrCode, iStrName, iStrNode, iStrGrpFlag, iStryyyymm
	Dim iIntRow

	On Error Resume Next
	' Add the top level(uniERP)
	With frm1
		Set iObjDummyNode = .uniTree1.Nodes.Add(, tvwChild, lgStrRootKey, lgStrRootDesc, C_Root, C_Root)
		
		For iIntRow = 1 To .vspdDataH.MaxRows
		
			.vspdDataH.Row = iIntRow
			
			.vspdDataH.Col = lgKeyPos(4)		' 코드명 
			iStrName = Trim(.vspdDataH.Text)
			.vspdDataH.Col = lgKeyPos(8)		' 코드 
			iStrCode = Trim(.vspdDataH.Text)
			.vspdDataH.Col = lgKeyPos(9)		' Parent
			iStrNode = Trim(.vspdDataH.Text)
			.vspdDataH.Col = lgKeyPos(10)		' Sales Group Flag
			iStrGrpFlag = Trim(.vspdDataH.Text)
			.vspdDataH.Col = lgKeyPos(11)		' Value '190001' means total amt. 
			iStryyyymm = Trim(.vspdDataH.Text)
			
			If iStryyyymm <> "190001" then exit for

			If iStrCode <> lgStrRootKey THEN
				If iStrGrpFlag = "N" Then
					Set iObjDummyNode = .uniTree1.Nodes.Add (iStrNode, tvwChild, iStrCode, "[" & Left(iStrCode,4) & "]" & iStrName, C_ORG)
				Else
					Set iObjDummyNode = .uniTree1.Nodes.Add (iStrNode, tvwChild, iStrCode, "[" & Left(iStrCode,4) & "]" & iStrName, C_GRP)
				End If
			Else
				.uniTree1.Nodes(iStrCode).Text = .uniTree1.Nodes(iStrCode).Text
			End If
		Next

		If Not(.uniTree1.Nodes(lgStrRootKey).Child Is Nothing) Then
			.uniTree1.Nodes(lgStrRootKey).Child.EnsureVisible						' Expand Tree	
		End If
		.uniTree1.Nodes(lgStrRootKey).Selected = True
	End With
End sub

'==========================================
Sub CopyVspdDataHToVspdDataH2()
	On Error Resume Next
	With frm1
		.vspdDataH2.MaxRows = 0
		
		.vspdDataH.col = 1
		.vspdDataH.col2 = .vspdDataH.MaxCols
		.vspdDataH.Row = 1
		.vspdDataH.Row2 = .vspdDataH.MaxRows
		
		' Dispay Total
		.vspdDataH2.MaxRows = .vspdDataH.MaxRows
		.vspdDataH2.Col = 1
		.vspdDataH2.Col2 = .vspdDataH2.MaxCols
		.vspdDataH2.Row = 1
		.vspdDataH2.Row2 = .vspdDataH2.MaxRows
		
		.vspdDataH2.Clip = .vspdDataH.Clip
	End With
	
End Sub

'==========================================
Sub CopyVspdDataHToVspdData(ByVal pvStrCode)
	Dim iIntRowForTotal, iIntStartRow, iIntCopyRows
	
	iIntCopyRows = 0
			
	frm1.vspdData.MaxRows = 0
	frm1.vspdData.Redraw = False

	With frm1.vspdDataH
		.col = lgKeyPos(8)
		.row = 1
		While(.Text <> pvStrCode)
			.row = .row + 1
		Wend
		
		'iIntRowForTotal = frm1.vspdDataH.SearchCol(lgKeyPos(8), 0, -1, pvStrCode, 0)
		iIntRowForTotal = .row

		.col = 1
		.col2 = .MaxCols
		.Row = iIntRowForTotal
		.Row2 = iIntRowForTotal
		
		' Dispay Total
		frm1.vspdData.MaxRows = 1
		frm1.vspdData.Col = 1
		frm1.vspdData.Col2 = frm1.vspdData.MaxCols
		frm1.vspdData.Row = 1
		frm1.vspdData.Row2 = 1
		frm1.vspdData.Clip = .Clip		
		
		' Display total for sub org.
		If Right(pvStrCode, 1) = C_GRP_SUFFIX Then
			iIntStartRow = 0
		Else
			.col = lgKeyPos(9)
			.row = 1
			While(.Text <> pvStrCode)
				.row = .row + 1
			Wend

			' iIntStartRow = frm1.vspdDataH.SearchCol(lgKeyPos(9), 0, -1, pvStrCode, 0)
			iIntStartRow = .row
		End If		
		
		If iIntStartRow > 0 Then
			.Row = iIntStartRow
			.Col = lgKeyPos(11)
		
			Do
				.Row = .Row + 1
				iIntCopyrows = iIntCopyrows + 1
				
				.Col = lgKeyPos(9)
				If .Text <> pvStrCode Then Exit Do
				.Col = lgKeyPos(11)
			Loop Until (.Text <> "190001")
			.col = 1
			.col2 = .MaxCols
			.Row = iIntStartRow
			.Row2 = iIntStartRow + iIntCopyRows - 1
			
			' Insert Rows
			frm1.vspdData.MaxRows = frm1.vspdData.MaxRows + iIntCopyRows

			frm1.vspdData.Col = 1
			frm1.vspdData.Col2 = frm1.vspdData.MaxCols
			frm1.vspdData.Row = 2
			frm1.vspdData.Row2 = frm1.vspdData.MaxRows
				
			frm1.vspdData.Clip = .Clip		
		End If

	End With
	
	frm1.vspdData.Row = 1 :	frm1.vspdDataH.Row = iIntRowForTotal
	frm1.vspdData.col = lgKeyPos(3) : frm1.vspdDataH.Col = lgKeyPos(1)
	frm1.vspdData.Text = frm1.vspdDataH.Text
	frm1.vspdData.col = lgKeyPos(4) : frm1.vspdDataH.Col = lgKeyPos(2)
	frm1.vspdData.Text = frm1.vspdDataH.Text
	
	frm1.vspdData.Redraw = True

End Sub

'==========================================
Sub CopyVspdDataH2ToVspdData2(ByVal pvStrCode)
	Dim iIntStartRow, iIntCopyRows

	iIntCopyRows = 0
			
	frm1.vspdData2.MaxRows = 0
	frm1.vspdData2.Redraw = False
	
	With frm1.vspdDataH2
		.col = lgKeyPos(8)
		.row = 1
		While(.Text <> pvStrCode)
			.row = .row + 1
		Wend
		iIntStartRow = .row
		'iIntStartRow = frm1.vspdDataH2.SearchCol(lgKeyPos(8), 0, -1, pvStrCode, 0)
		
		.Row = iIntStartRow
		.Col = lgKeyPos(8)
		
		Do
			.Row = .Row + 1
			iIntCopyrows = iIntCopyrows + 1
		Loop Until (.Text <> pvStrCode)

		.col = 1
		.col2 = 7
		.Row = iIntStartRow
		.Row2 = iIntStartRow + iIntCopyRows - 1
		
		frm1.vspdData2.MaxRows = iIntCopyRows

		frm1.vspdData2.Col = 1
		frm1.vspdData2.Col2 = frm1.vspdData2.MaxCols
		frm1.vspdData2.Row = 1
		frm1.vspdData2.Row2 = iIntCopyRows
		
		frm1.vspdData2.Clip = .Clip		
		
	End With
	
	frm1.vspdData2.Redraw = True
End Sub

'==========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	Case C_PopSalesOrg												
		iArrParam(1) = "B_SALES_ORG "						
		iArrParam(2) = Trim(frm1.txtSalesOrg.value)			
		iArrParam(3) = ""									
		iArrParam(4) = "LVL = " & frm1.cboSalesOrgLvl.value	
		iArrParam(5) = "영업조직"							
			
		iArrField(0) = "SALES_ORG"							
		iArrField(1) = "SALES_ORG_NM"								    
		iArrHeader(0) = "영업조직"						
		iArrHeader(1) = "영업조직명"						

		frm1.txtSalesOrg.focus
	End Select
 
	iArrParam(0) = iArrParam(5)							

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True
	End If	
	
End Function

'==========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopSalesOrg
		frm1.txtSalesOrg.value = pvArrRet(0) 
		frm1.txtSalesOrgNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'========================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs("minor_cd, minor_nm ","b_minor","major_cd = " & FilterVar("S0016", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboSalesOrgLvl, lgF0, lgF1, Chr(11))
	Call SetCombo(frm1.cboQueryData, "B", "매출")
	Call SetCombo(frm1.cboQueryData, "T", "세금계산서")
End Sub

'========================================
Function txtSalesOrg_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtSalesOrg.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", .cboSalesOrgLvl.value, "" & FilterVar("SO", "''", "S") & "", C_PopSalesOrg) Then
				.txtSalesOrg.value = ""
				.txtSalesOrgNm.value = ""
				.txtSalesOrg.focus
			Else
				.txtFromDt.focus
			End If
			txtSalesOrg_OnChange = False
		Else
			.txtSalesOrgNm.value = ""
		End If
	End With
	lgBlnSalesOrgChg = False
End Function

'==========================================
Function cboSalesOrgLvl_OnChange()
	With frm1
		.txtSalesOrg.value = ""
		.txtSalesOrgNm.value = ""
		If .cboSalesOrgLvl.value = "" Then
			ggoOper.SetReqAttr .txtSalesOrg , "Q"
			.btnSalesOrg.disabled = True
		Else
			ggoOper.SetReqAttr .txtSalesOrg , "D"
			.btnSalesOrg.disabled = False
		End If
	End With
End Function

'========================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
	End If
End Sub

'==========================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
	End If
End Sub

'==========================================
Sub txtFromDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'==========================================
Sub txtToDt_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'==========================================
Sub txtSalesOrg_OnKeyDown()
	lgBlnSalesOrgChg = True
	lgBlnFlgConChg = True
End Sub

'========================================= 
Sub uniTree1_onAddImgReady()
    Call SetToolbar("11000000000011")										
End Sub

'==========================================
Sub uniTree1_NodeClick(pvObjNode)
	If pvObjNode.Key = lgStrPrevNodeKey Then Exit Sub
	Call CopyVspdDataHToVspdData(pvObjNode.Key)
	lgStrPrevNodeKey = pvObjNode.Key
	
	With frm1.vspdData
		If .MaxRows > 0 Then			
			.Row = 1
			.Col = lgKeyPos(8)
			Call CopyVspdDataH2ToVspdData2(.Text)
		End If
	End With
End Sub

'========================================
Sub vspdData_Click( Col,  Row)
	gMouseClickStatus = "SPC"

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort Col, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort Col, lgSortKey_A
            lgSortKey_A = 1
        End If    
        Exit Sub
    End If
    
  	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
End Sub

'========================================
Sub vspdData2_Click( Col,  Row)
    Dim ii
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort Col, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort Col, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If

	gMouseClickStatus = "SP2C"
End Sub

'==========================================
Sub vspdData_LeaveRow(ByVal pvIntRow, ByVal pvBlnRowWasLast, ByVal pvBlnRowChanged, ByVal pvBlnAllCellsHaveData, ByVal pvIntNewRow, ByVal pvIntNewRowIsLast, pvBlnCancel) 
	With frm1.vspdData
		.Row = pvIntNewRow
		.Col = lgKeyPos(8)
		Call CopyVspdDataH2ToVspdData2(.Text)
	End With
End Sub

'==========================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>영업조직별월매출현황(T)</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS="TD5" NOWRAP>매출채권일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFromDt" CLASS="FPDTYYYYMMDD" tag="12X1" Alt="시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToDt" CLASS="FPDTYYYYMMDD" tag="12X1" Alt="종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS="TD5" NOWRAP>조회기준</TD>
	                        		<TD CLASS="TD6" NOWRAP>
                						<SELECT Name="cboQueryData" ALT="조회기준" CLASS ="cbonormal" tag="12"><OPTION></OPTION></SELECT>
		                    		</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업조직레벨</TD>
	                        		<TD CLASS="TD6" NOWRAP>
                						<SELECT Name="cboSalesOrgLvl" ALT="영업조직레벨" CLASS ="cbonormal" tag="12"><OPTION></OPTION></SELECT>
		                    		</TD>
									<TD CLASS=TD5 NOWRAP>영업조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesOrg" ALT="영업조직" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnSalesOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesOrg">&nbsp;<INPUT NAME="txtSalesOrgNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>B/L포함여부</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBLFlg" TAG="11X" VALUE="A" ID="rdoBLFlgA"><LABEL FOR="rdoBLFlgA">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBLFlg" TAG="11X" VALUE="Y" ID="rdoBLFlgY"><LABEL FOR="rdoBLFlgY">예</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoBLFlg" TAG="11X" VALUE="N" CHECKED ID="rdoBLFlgN"><LABEL FOR="rdoBLFlgN">아니오</LABEL>			
									</TD>
									<TD CLASS=TD5 NOWRAP></TD> 
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD WIDTH = 30%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=uniTree1 width=100% height=100% <%=UNI2KTV_IDVER%>> <PARAM NAME="ImageWidth" VALUE="16">  <PARAM NAME="ImageHeight" VALUE="16">  <PARAM NAME="LineStyle" VALUE="1"> <PARAM NAME="Style" VALUE="7">  <PARAM NAME="LabelEdit" VALUE="1">  </OBJECT>');</SCRIPT>
								</TD>
								<TD WIDTH=*>
									<TABLE <%=LR_SPACE_TYPE_20%>>
										<TR HEIGHT="50%">
											<TD WIDTH="100%" colspan=4>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> name=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT4> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
										</TR>
										<TR HEIGHT="50%">
											<TD WIDTH="100%" colspan=4>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT5> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
										</TR>
									</TABLE>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdDataH WIDTH=0 HEIGHT=0 TAG="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdDataH2 WIDTH=0 HEIGHT=0 TAG="2" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
		</TD>
	</TR>
</TABLE>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
