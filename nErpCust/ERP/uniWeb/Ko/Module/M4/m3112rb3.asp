<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1          '☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   Dim iTotstrData
   
   Dim strItemNm
   Dim iFrPoint
   iFrPoint=0

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint		= C_SHEETMAXROWS_D * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    
    SetConditionData = FALSE
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strItemNm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtItem")) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			Exit Function
		End If
	End If   	

	SetConditionData = TRUE
	
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(0)
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(1,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "M3112QA001" 										' main query(spread sheet에 뿌려지는 query statement)
     UNISqlId(1) = "s0000qa001"											' 품목명 
     

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	strVal = " "

	If Len(Request("txtItem")) Then			'품목 
		strVal = " AND B.ITEM_CD = " & FilterVar(Trim(UCase(Request("txtItem"))), "''" , "S") & " "
	End If
	arrVal(0)=FilterVar(Trim(UCase(Request("txtItem"))), "''" , "S")

	If Len(Request("txtPurgrp")) Then		'구매그룹 
		strVal = strVal & " AND A.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtPurgrp"))), "''" , "S") & " "
	End If
	
	If Len(Request("txtBeneficiary")) Then	'수혜자 
		strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), "''" , "S") & " "
	End If
	
	If Len(Request("txtCurrency")) Then		'화폐 
		strVal = strVal & " AND A.PO_CUR = " & FilterVar(Trim(UCase(Request("txtCurrency"))), "''" , "S") & " "
	End If
	
'	If Len(Request("txtPayTerms")) Then		'결제방법 
'		strVal = strVal & " AND A.PAY_METH ='" & Trim(Request("txtPayTerms")) & "'"
'	End If
	
	If Len(Request("txtPONo")) Then			'발주번호 
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Trim(UCase(Request("txtPONo"))), "''" , "S") & " "
	End If
	
	'2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), "''" , "S") & "  "		
	End If

    UNIValue(0,1) = strVal 
    UNIValue(1,0) = arrVal(0)
'    UNIValue(0,UBound(UNIValue,2)) = "ORDER BY A.PO_NO DESC,B.PO_SEQ_NO ASC"
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))	
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
	IF SetConditionData() = FALSE THEN EXIT SUB
 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>
<Script Language=vbscript>
With parent
    .frm1.txtItemNm.Value	= "<%=ConvSPChars(strItemNm)%>"

    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.txtHItem.value			= "<%=ConvSPChars(Request("txtItem"))%>"
			.frm1.txtHPurgrp.value		= "<%=ConvSPChars(Request("txtPurgrp"))%>"
			.frm1.txtHBeneficiary.value	= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
			.frm1.txtHCurrency.value		= "<%=ConvSPChars(Request("txtCurrency"))%>"
			'.frm1.txtHPayTerms.value		= "<%=ConvSPChars(Request("txtPayTerms"))%>"
			.frm1.txtHPONo.value			= "<%=ConvSPChars(Request("txtPONo"))%>"
       End If
       
       .ggoSpread.Source  = parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       .ggoSpread.SSShowData "<%=iTotstrData%>","F"          '☜ : Display data

		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",6),"C","I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",7),"A","I","X","X")

       .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       .DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   
End With    
</Script>	
