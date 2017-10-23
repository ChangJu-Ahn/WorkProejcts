<%'======================================================================================================
'*  1. Module Name          : 구매 
'*  2. Function Name        : 매입관리 
'*  3. Program ID           : m5111pa1
'*  4. Program Name         : 매입번호 
'*  5. Program Desc         : 매입내역등록의 매입번호 
'*  6. Comproxy List        : M51118ListIvHdrSvr
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : Shin Jin Hyen	
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date 표준적용 
'*							  2002/05/08 ADO 변환 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3 '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Const SortNo=2											  ' Sort 종류 
Dim iTotstrData

Dim strIvType											  ' 매입형태명 
Dim strPurGrp											  ' 구매그룹명 
Dim strSupplier											  ' 공급처명 

Dim iPrevEndRow
Dim iEndRow
   
    Call HideStatusWnd
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB") 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
	iPrevEndRow = 0
    iEndRow = 0
        
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
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
Function SetConditionData()
    'On Error Resume Next
    
    SetConditionData = false
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strIvType =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtIvType")) Then
			Call DisplayMsgBox("970000", vbInformation, "매입형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If   	
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strPurGrp =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtPur_grp")) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If			
    End If     
     
	If Not(rs3.EOF Or rs3.BOF) Then
        strSupplier =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtSupplier")) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If			
    End If     
 
    SetConditionData = True
  
End Function 

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,2)

    UNISqlId(0) = "M5111QA001"
    UNISqlId(1) = "M5111QA103"					'매입형태명    
    UNISqlId(2) = "s0000qa019"					'구매그룹명    
    UNISqlId(3) = "s0000qa002"					'공급처명    

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    

	strVal = " "

	If Len(Request("txtSupplier")) Then	
		strVal = strVal & "AND A.BP_CD =  " & FilterVar(Request("txtSupplier"), "''", "S") & " " & CHR(13)
	end if
		arrVal(2) = FilterVar(Trim(Request("txtSupplier")), "", "S")	

	If Len(Request("txtGroup")) Then		
		strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Request("txtGroup"), "''", "S") & " " & CHR(13)
	End If
		arrVal(1) = FilterVar(Trim(Request("txtGroup")), "", "S")
		
	If Len(Request("txtIvType")) Then
		strVal = strVal & " AND A.IV_TYPE_CD =  " & FilterVar(Request("txtIvType"), "''", "S") & " " & CHR(13)
	End If	
	arrVal(0) = FilterVar(Trim(Request("txtIvType")), "", "S")

    If Len(Trim(Request("txtFrIvDt"))) Then
		strVal = strVal & " AND A.IV_DT >= " & FilterVar(UNIConvDate(Request("txtFrIvDt")), "''", "S") & " " & CHR(13)
	End If		

    If Len(Trim(Request("txtToIvDt"))) Then
		strVal = strVal & " AND A.IV_DT <= " & FilterVar(UNIConvDate(Request("txtToIvDt")), "''", "S") & " " & CHR(13)
	End If		

	If Len(Request("txtRadio")) Then
		strVal = strVal & " AND A.POSTED_FLG = " & FilterVar(Trim(Request("txtRadio")), "''", "S") & " " & CHR(13)
	End If		
	
	UNIValue(0,1) = strVal   
    UNIValue(1,0) = arrVal(0)						'매입형태명 
    UNIValue(2,0) = arrVal(1)									'구매그룹명    
    UNIValue(3,0) = arrVal(2)									'공급처명    
    

    UNIValue(0,UBound(UNIValue,2)) = "ORDER BY A.IV_NO DESC"
    UNILock = DISCONNREAD :	UNIFlag = "1"                       '☜: set ADO read mode
 
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
    
    If SetConditionData = False Then Exit sub
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>
<Script Language=vbscript>
    With parent
		.frm1.txtIvTypeNm.Value 	= "<%=ConvSPChars(strIvType)%>"
		.frm1.txtGroupNm.Value 		= "<%=ConvSPChars(strPurGrp)%>"
		.frm1.txtSupplierNm.Value 	= "<%=ConvSPChars(strSupplier)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnIvType.Value 	= "<%=ConvSPChars(Request("txtIvType"))%>"
				.frm1.hdnSupplier.Value = "<%=ConvSPChars(Request("txtSupplier"))%>"
				.frm1.hdnGroup.Value 	= "<%=ConvSPChars(Request("txtGroup"))%>"
				.frm1.hdnFrDt.Value 	= "<%=Request("txtFrIvDt")%>"
				.frm1.hdnToDt.Value 	= "<%=Request("txtToIvDt")%>"
			End If    
			       
			.ggoSpread.Source    = .frm1.vspdData 
			Parent.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=iTotstrData%>", "F"          '☜ : Display data
		   
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,Parent.GetKeyPos("A",8), Parent.GetKeyPos("A",7),"A", "I" ,"X","X")
       
			.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
			.DbQueryOk
			Parent.frm1.vspdData.Redraw = True
		End If
	End with
</Script>	
