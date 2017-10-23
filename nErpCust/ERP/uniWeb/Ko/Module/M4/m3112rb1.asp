<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3112rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : P/O 내역참조 PopUp Transaction 처리용 ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/07/11																*
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2002/05/06 : ADO Conv.												*
'********************************************************************************************************
%>
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1              '☜ : DBAgent Parameter 선언 
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
	
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(1,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "M3112RA101" 										' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(1) = "M3112RA102" 

    strVal = " "
    
	IF Len(Trim(Request("txtPoNo"))) THEN
		strVal = " AND 	lhdr.PO_NO = " & FilterVar(Trim(UCase(Request("txtPoNo"))), " " , "S") & "  "
	END IF
    
	IF Len(Trim(Request("txtBeneficiaryCd"))) THEN
		strVal = strVal & " AND 	lhdr.BP_CD = " & FilterVar(Trim(UCase(Request("txtBeneficiaryCd"))), " " , "S") & "  "
	END IF
	
	IF Len(Trim(Request("txtItemCd"))) THEN
		strVal = strVal & " AND mdtl.ITEM_CD = " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & "  "
	END IF
	
	IF Len(Trim(Request("txtGrpCd"))) THEN
		strVal = strVal & " AND lhdr.PUR_GRP = " & FilterVar(Trim(UCase(Request("txtGrpCd"))), " " , "S") & "  "
	END IF
	
	IF Len(Trim(Request("txtCurrency"))) THEN
		strVal = strVal & " AND lhdr.PO_CUR = " & FilterVar(Trim(UCase(Request("txtCurrency"))), " " , "S") & "  "
	END IF
	
	'IF Len(trim(Request("txtPayMethCd"))) THEN
	'	strVal = strVal & " AND lhdr.PAY_METH ='" & FilterVar(Trim(UCase(Request("txtPayMethCd"))), " " , "SNM") & "' "
	'END IF
	
	IF Len(Trim(Request("txtIncotermsCd"))) THEN
		strVal = strVal & " AND lhdr.INCOTERMS = " & FilterVar(Trim(UCase(Request("txtIncotermsCd"))), " " , "S") & "  "
	END IF
	
	'---2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND mdtl.TRACKING_NO = " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & "  "		
	End If

    UNIValue(0,0) = Trim(lgSelectList)		                            '☜: Select 절에서 Summary    필드 
    UNIValue(0,1) = strVal   
    
    '--- 2004-08-19 by Byun Jee Hyun for UNI Code
    UNIValue(1,0) =  FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") 				'품목 
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    
    
	'UNIValue(0,UBound(UNIValue,2)) = " ORDER BY MDTL.ITEM_CD DESC "
    'UNIValue(0,UBound(UNIValue,2)) = " ORDER BY MDTL.PO_SEQ_NO ASC "
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
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  	
   
End Sub

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
        strItemNm = rs1(1)
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtItemCd"))) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			EXIT FUNCTION
		End If
	End If 
	
	SetConditionData = TRUE
	
End Function

%>
<Script Language=vbscript>
with parent
    .frm1.txtItemNm.value = "<%=ConvSPChars(strItemNm)%>" 

    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.hdnPoNo.value				= "<%=ConvSPChars(Request("txtPONo"))%>"
			.frm1.hdnLcFlg.value			= "<%=ConvSPChars(Request("txtLcFlg"))%>"
			.frm1.hdnItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
			.frm1.hdnGrpCd.value			= "<%=ConvSPChars(Request("txtGrpCd"))%>"
			.frm1.hdnBeneficiaryCd.value	= "<%=ConvSPChars(Request("txtBeneficiaryCd"))%>"
			.frm1.hdnCurrency.value			= "<%=ConvSPChars(Request("txtCurrency"))%>"
			'.frm1.hdnPayMethCd.value		= "<%=ConvSPChars(Request("txtPayMethCd"))%>"
			.frm1.hdnIncotermsCd.value		= "<%=ConvSPChars(Request("txtIncotermsCd"))%>"
       End If
			.ggoSpread.Source  = .frm1.vspdData
			Parent.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=iTotstrData%>","F"          '☜ : Display data

			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",6),"C","I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",7),"A","I","X","X")
			
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
			.DbQueryOk
			Parent.frm1.vspdData.Redraw = True
    End If  
end with     
</Script>	

