<% Option explicit%>
<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3212rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 발주내역참조 PopUp Transaction 처리용 ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     : Sun-jung Lee
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************
%>
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next													   '실행 오류가 발생할 때 오류가 발생한 문장 바로 다음에 실행이 계속될 수 있는 문으로 컨트롤을 옮길 수 있도록 지정합니다.				
Err.Clear

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4                '☜ : DBAgent Parameter 선언 
	Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
	Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iTotstrData
	
    Dim iPrevEndRow
    Dim iEndRow
   
	Dim strBeneficiaryNm
	Dim strPurGrpNm
	Dim strPaymeth
	Dim strIncoterms
	
	Dim strPlantNm
	Dim strItemNm
	Dim strVatTypeNm 	
	
	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
	
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
	iPrevEndRow = 0
    iEndRow = 0

	SELECT CASE REQUEST("txtMode")								 '☜ : onChange 에서 호출할경우와 메인쿼리인경우 
		CASE "changeItemPlant"
			Call FixUNISQLData2()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
			Call QueryData2()										 '☜ : DB-Agent를 통한 ADO query
	
		CASE ELSE
			Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
			Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
	END SELECT

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(3)
	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
    
    UNISqlId(0) = "M4111RA101"
    UNISqlId(1) = "s0000qa009"											'공장 
    UNISqlId(2) = "s0000qa016"											'품목 
    UNISqlId(3) = "s0000qa026"											'VAT명     

    UNIValue(0,0) = Trim(lgSelectList)		                            '☜: Select 절에서 Summary    필드 

	strVal = ""
  
  	If Len(Request("txtMvmtNo")) Then
		strVal = strVal & " AND A.MVMT_RCPT_NO = " & FilterVar(Request("txtMvmtNo"), "''", "S") & " "
	End If
	
	'---2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & "  "		
	End If

	If Len(Request("txtPlantCd")) Then
		strVal = strVal & " AND B.PLANT_CD = " & FilterVar(Request("txtPlantCd"), "''", "S") & " "
	End If
	arrVal(0) = FilterVar(Trim(Request("txtPlantCd")), "", "S")
	
	If Len(Request("txtItemcd")) Then
		strVal = strVal & " AND C.ITEM_CD = " & FilterVar(Request("txtItemcd"), "''", "S") & " "
	End If
	arrVal(1) = FilterVar(Trim(Request("txtItemCd")), "", "S")

	If Len(Request("txtVatType")) Then
		strVal = strVal & " AND i.minor_cd = " & FilterVar(Request("txtVatType"), "''", "S") & " "
	End If
	arrVal(2) = FilterVar(Trim(Request("txtVatType")), "", "S")

    If Len(Trim(Request("txtFrDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT >= " & FilterVar(UNIConvDate(Request("txtFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <= " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""		
	End If

	If Len(Trim(Request("txtPoNo"))) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Request("txtPoNo"), "''", "S") & " "		
	End If
	
	If Len(Trim(Request("txtSppl"))) Then
		strVal = strVal & " AND F.BP_CD = " & FilterVar(Request("txtSppl"), "''", "S") & " "		
	End If
	
	If Len(Trim(Request("txtIvType"))) Then
		strVal = strVal & " AND F.IV_TYPE = " & FilterVar(Request("txtIvType"), "''", "S") & " "		
	End If
	
	'2009-09-02 화폐단위에 상관없이 불러오게 수정... 김지한 과장 요청
	'If Len(Trim(Request("txtPoCur"))) Then
	'	strVal = strVal & " AND F.PO_CUR = " & FilterVar(Request("txtPoCur"), "''", "S") & " "		
	'End If
	
	If UCase(Trim(Request("txtLcKind"))) <> "N" Then
		'LC번호가 있는경우(Local LC 후 입출고 참조하는 경우)
		strVal = strVal & " AND L.PAY_METHOD = " & FilterVar(Request("txtPayMeth"), "''", "S") & " "		
    End If
    
    '매입내역의 입고참조시 매입일자 이전의 입출고 건만 조회되도록 변경(2005-10-28)
    If Len(Trim(Request("txtIvDt"))) Then
		strVal = strVal & " AND A.MVMT_DT <= " & FilterVar(UNIConvDate(Request("txtIvDt")), "''", "S") & ""		
	End If		

     If Request("gPlant") <> "" Then
        strVal = strVal & " AND a.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND a.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND a.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND a.MVMT_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   

	
    UNIValue(0,1) = strVal
    UNIValue(1,0) = arrVal(0)  	'공장 
    UNIValue(2,0) = arrVal(1)  	'품목 
    UNIValue(3,0) = arrVal(2)	'VAT
   
    'UNIValue(0,UBound(UNIValue,2)) = ""
    UNIValue(0,UBound(UNIValue,2)) = " ORDER BY A.MVMT_RCPT_NO DESC,A.PO_NO DESC,A.PO_SEQ_NO ASC "			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
                        '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 (ONCHAGE 에서 호출)
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData2()

    Dim strVal
	Dim arrVal(2)
	Redim UNISqlId(2)                                                   '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "s0000qa009"											'공장 
    UNISqlId(1) = "s0000qa016"											'품목 
    UNISqlId(2) = "s0000qa026"											'VAT명     

    UNIValue(0,0) = Trim(lgSelectList)		                            '☜: Select 절에서 Summary    필드 

	strVal = " "
	
	'If Len(Request("txtPlantCd")) Then
		arrVal(0) = FilterVar(Trim(Request("txtPlantCd")), "", "S")
	'End If
	
	'If Len(Request("txtItemcd")) Then
		arrVal(1) = FilterVar(Trim(Request("txtItemCd")), "", "S")
	'End If

	'If Len(Request("txtVatType")) Then
		arrVal(2) = FilterVar(Trim(Request("txtVatType")), "", "S")
	'End If

    UNIValue(0,0) = arrVal(0)  	'공장 
    UNIValue(1,0) = arrVal(1)  	'품목 
    UNIValue(2,0) = arrVal(2)	'VAT
   
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
                        '☜: set ADO read mode
	
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If SetConditionData = False Then Exit Sub 
         
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
   
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data2
'----------------------------------------------------------------------------------------------------------
Sub QueryData2()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    Call  SetConditionData()
End Sub
    
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
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
    On Error Resume Next
    
    SetConditionData = false 
     
	If Not(rs1.EOF Or rs1.BOF) Then
        strPlantNm = rs1(1)
        Set rs1 = Nothing
	else
	    Set rs1 = Nothing
		If Len(Request("txtPlantCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit function
		End If
	End If  
	

	If Not(rs2.EOF Or rs2.BOF) Then
        strItemNm = rs2(1)
        Set rs2 = Nothing	
	else
	    Set rs2 = Nothing
		If Len(Request("txtItemcd")) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit function
		End If

	End If  

	If Not(rs3.EOF Or rs3.BOF) Then
        strVatTypeNm = rs3(1)
        Set rs3 = Nothing
	else
	    Set rs3 = Nothing
		If Len(Request("txtVatType")) Then
			Call DisplayMsgBox("970000", vbInformation, "VAT유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit function
		End If
    End If  

	
	SetConditionData = True
	
End Function

%>

<Script Language=vbscript>
parent.frm1.txtPlantNm.Value 		= "<%=ConvSPChars(strPlantNm)%>"
parent.frm1.txtItemNm.Value 		= "<%=ConvSPChars(strItemNm)%>"
parent.frm1.txtVatNm.Value 			= "<%=ConvSPChars(strVatTypeNm)%>"	

    With parent
		If "<%=lgDataExist%>" = "Yes" Then
		   If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.hdnPoNo.value = "<%=ConvSPChars(request("txtPoNo"))%>"
			.frm1.hdnFrMvmtDt.value = "<%=request("txtFrDt")%>"
			.frm1.hdnToMvmtDt.value = "<%=request("txtToDt")%>"
			.frm1.hdnSupplierCd.value = "<%=ConvSPChars(request("txtSppl"))%>"
			.frm1.hdnMvmtNo.value = "<%=ConvSPChars(request("txtMvmtNo"))%>"
			.frm1.hdnRefType.value = "<%=ConvSPChars(request("txtRefType"))%>"
			.frm1.hdnIvType.value = "<%=ConvSPChars(request("txtIvType"))%>"
			.frm1.hdnPoCur.value = "<%=ConvSPChars(request("txtPoCur"))%>"
			.frm1.hdnPlantCd.value = "<%=ConvSPChars(request("txtPlantCd"))%>"
			.frm1.hdnItemCd.value = "<%=ConvSPChars(request("txtItemCd"))%>"
			.frm1.hdnVatType.value = "<%=ConvSPChars(request("txtVatType"))%>"		   
		   End If
		   
		   .ggoSpread.Source  = .frm1.vspdData
		   Parent.frm1.vspdData.Redraw = false
		   .ggoSpread.SSShowData "<%=iTotstrData%>", "F"          '☜ : Display data
		   
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",12),"D", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",19),"C", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",20),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",25),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",26),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",29),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",30),"A", "I" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",36),"C", "I" ,"X","X")
       
		   .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
		   .DbQueryOk
		   Parent.frm1.vspdData.Redraw = True
		End If  
	End with
</Script>	
 	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

