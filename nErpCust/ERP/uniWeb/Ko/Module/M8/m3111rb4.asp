<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3111rb4.asp																*
'*  4. Program Name         : 발주참조(매입세금계산서)				   					    	*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2000/05/10																*
'*  9. Modifier (First)     : Shin jin hyun																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************

On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           '☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   Dim iTotstrData
   
   Dim strPOTypeName
   Dim strBPName
   Dim strPGRName
   
   Dim iPrevEndRow
   Dim iEndRow
    
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"
	iPrevEndRow = 0
    iEndRow = 0
    
    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim PvArr
    
    Const C_SHEETMAXROWS_D = 100    
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                 
    End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
            PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
function SetConditionData()
    'On Error Resume Next
	SetConditionData = TRUE


    If Not(rs1.EOF Or rs1.BOF) Then
        strPOTypeName = rs1(0)
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtPotype"))) Then
			Call DisplayMsgBox("970000", vbInformation, "발주형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			SetConditionData = FALSE
			exit function		
		End If
	End If 
	
	If Not(rs2.EOF Or rs2.BOF) Then
        strBPName = rs2(1)
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtSupplier"))) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			SetConditionData = FALSE
			exit function
		End If
	End If  
	
	If Not(rs3.EOF Or rs3.BOF) Then
       strPGRName = rs3(0)
   		Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "발주담당", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			SetConditionData = FALSE
			exit function
		End If
	End If  
	
End function
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
	Dim arrVal(3)                                                                          
     UNISqlId(0) = "m3111ra401"
     UNISqlId(1) = "s0000qa020" 
     UNISqlId(2) = "s0000qa002" 
     UNISqlId(3) = "s0000qa022"                 
     
     '--- 2004-08-19 by Byun Jee Hyun for UNICODE
     UNIValue(0,0) = lgSelectList		                              '☜: Select 절에서 Summary    필드 

	strVal = " "

	If Len(Request("txtPotype")) Then
		strVal = " AND B.PO_TYPE_CD = " & FilterVar(Request("txtPotype"), "''", "S") & " "
	Else
		strVal = ""
	End If
	arrVal(1) =  FilterVar(Trim(Request("txtPotype")), "", "S")
	
	If Len(Request("txtSupplier")) Then
		strVal = strVal & " AND B.BP_CD = " & FilterVar(Request("txtSupplier"), "''", "S") & " "
	Else
		strVal = strVal & ""
	End If
	arrVal(2) =  FilterVar(Trim(Request("txtSupplier")), "", "S")	
	
	If Len(Trim(Request("txtFrPoDt"))) Then
		strVal = strVal & " AND B.PO_DT >= " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & ""
	Else
		strVal = strVal & " AND B.PO_DT >=" & "" & FilterVar("1900/01/01", "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtToPoDt"))) Then
		strVal = strVal & " AND B.PO_DT <= " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & ""		
	Else
		strVal = strVal & " AND B.PO_DT <=" & "" & FilterVar("2999/12/31", "''", "S") & ""		
	End If

	If Len(Request("txtGroup")) Then
		strVal = strVal & " AND B.PUR_GRP = " & FilterVar(Request("txtGroup"), "''", "S") & " "
	Else
		strVal = strVal & ""
	End If
	arrVal(3) = FilterVar(Trim(Request("txtGroup")),"", "S")

    UNIValue(0,1) = strVal
    UNIValue(1,0) = arrVal(1)   
    UNIValue(2,0) = arrVal(2)
    UNIValue(3,0) = arrVal(3)               
       
    UNIValue(0,UBound(UNIValue,2)) = "ORDER BY B.PO_NO DESC"
    'UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    if SetConditionData = false then exit sub
 
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
parent.frm1.txtPotypeNm.value = "<%=ConvSPChars(strPOTypeName)%>"
parent.frm1.txtSupplierNm.value = "<%=ConvSPChars(strBPName)%>"
parent.frm1.txtGroupNm.value = "<%=ConvSPChars(strPGRName)%>"
    
    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			parent.frm1.hdnPotype.value	= "<%=ConvSPChars(Request("txtPotype"))%>"  
			parent.frm1.hdnSupplier.value = "<%=ConvSPChars(Request("txtSupplier"))%>" 
			parent.frm1.hdnFrDt.value	= "<%=ConvSPChars(Request("txtFrPoDt"))%>" 
			parent.frm1.hdnToDt.value	= "<%=ConvSPChars(Request("txtToPoDt"))%>" 
			parent.frm1.hdnGroup.value	= "<%=ConvSPChars(Request("txtGroup"))%>"
	
       End If

       parent.ggoSpread.Source  = parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       parent.ggoSpread.SSShowData "<%=iTotstrData%>", "F"          '☜ : Display data
		   
	   Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,Parent.GetKeyPos("A",5), Parent.GetKeyPos("A",4),"A", "I" ,"X","X")
       
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       parent.DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   
</Script>	
