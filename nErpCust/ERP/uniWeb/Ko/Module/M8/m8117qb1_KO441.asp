<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M8117QB101
'*  4. Program Name         : 
'*  5. Program Desc         : 매입가계정잔액현황조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/06/17
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Jin Ha
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%	
	Call HideStatusWnd
	call LoadBasisGlobalInf()
	call LoadInfTB19029B("Q", "M","NOCOOKIE","QB") 
	call LoadBNumericFormatB("Q","M","NOCOOKIE","QB")

    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                '☜ : DBAgent Parameter 선언 
    Dim lgTailList
    Dim lgPageNo
    Dim istrData
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
    Dim iPrevEndRow
    Dim iEndRow
    Dim iTotstrData
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	iPrevEndRow = 0
    iEndRow = 0
    lgOpModeCRUD  = Request("txtMode") 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call  SubBizQueryMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	Err.Clear
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)  
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	
	Call FixUNISQLData()
	Call QueryData()	
	
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal, strVal1
	Redim UNISqlId(2)														'☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,4)														'⊙: DB-Agent로 전송될 parameter를 위한 변수 
																			'parameter의 수에 따라 변경함 
    UNISqlId(0) = "M8117QA101"												'Detial 											
    UNISqlId(1) = "M8117QA102" 												'Summary
    UNISqlId(2) = "S0000QA023"												'입출고유형명 
	'1? : Iv_dt, 2? : mvmt_dt, 3? : io_type_cd, 4? :차이 or 전체 
	
	strVal = ""
	strVal1 = ""
	strVal = strVal & " a.mvmt_dt >=  " & FilterVar(UNIConvDate(Request("txtMvmtFrDt")), "''", "S") & " "
	strVal = strVal & " and a.mvmt_dt <=  " & FilterVar(UNIConvDate(Request("txtMvmtToDt")), "''", "S") & " "
	
	If Request("txtIvFrDt") <> "" Then
		strVal1 = strVal1 & " and n.iv_dt >=  " & FilterVar(UNIConvDate(Request("txtIvFrDt")), "''", "S") & " "
	End IF
	
	If Request("txtIvToDt") <> "" Then
		strVal1 = strVal1 & " and n.iv_dt <=  " & FilterVar(UNIConvDate(Request("txtIvToDt")), "''", "S") & " "
	End If
	
     If Request("gPlant") <> "" Then
        strVal1 = strVal1 & " AND m.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal1 = strVal1 & " AND n.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal1 = strVal1 & " AND n.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        strVal1 = strVal1 & " AND n.IV_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   
	
	UNIValue(0,0) = strVal1	
	UNIValue(1,0) = strVal1	
	UNIValue(0,1) = strVal	
	UNIValue(1,1) = strVal
		
	strVal  = ""
	strVal1 = ""
	
	If Trim(Len(Request("txtMvmtType"))) Then
		strVal = strVal & " and a.io_type_cd =  " & FilterVar(UCase(Request("txtMvmtType")), "''", "S") & "  "
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
	
	If Trim(Request("rdoAppflg")) = "A" Then
		strVal1 = strVal1 & " or (IsNull(sum(b.mvmt_loc_amt),0) - IsNull(sum(b.grir_amt),0) = 0 ) "
	End If
	
	
	UNIValue(0,2) = strVal
	UNIValue(0,3) = strVal1
	
	UNIValue(1,2) = strVal
	UNIValue(1,3) = strVal1
	
	UNIValue(2,0)  = FilterVar(Trim(UCase(Request("txtMvmtType"))), "''" , "S") 
	
	UNIValue(0,UBound(UNIValue,2)) = " Order by b.mvmt_dt Desc, b.mvmt_rcpt_no,b.io_type_cd "
	UNILock = DISCONNREAD :	UNIFlag = "1"                                
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    Dim iSumMvmtAmt
	Dim	iSumGrirAmt
	Dim iStrIoTypeNm
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    Dim FalsechkFlg
    
    FalsechkFlg = False    
	
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
    Else    
		iSumMvmtAmt = rs1(0)
		iSumGrirAmt = rs1(1)
		rs1.Close
        Set rs1 = Nothing
    End If
    
	If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtMvmtType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "입고유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		iStrIoTypeNm = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Response.End
    Else    
        Call  MakeSpreadSheetData()
    End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent"				& vbCr
	Response.Write "	.CurFormatNumericOCX"   & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
	Response.Write "    .frm1.vspdData.Redraw = False   "                      & vbCr   
    Response.Write "	.ggoSpread.SSShowData        """ & iTotstrData	    & """ ,""F""" & vbCr
   
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iPrevEndRow + 1 & "," & iEndRow & "	,.parent.gCurrency		,.C_MvmtAmt		 ,""A"" ,""Q"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iPrevEndRow + 1 & "," & iEndRow & "	,.parent.gCurrency		,.C_IvLocAmt	 ,""A"" ,""Q"",""X"",""X"")" & vbCr
    Response.Write  "    Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData," & iPrevEndRow + 1 & "," & iEndRow & "	,.parent.gCurrency		,.C_IvTmpAcctBal ,""A"" ,""Q"",""X"",""X"")" & vbCr
	
	Response.Write "	.lgPageNo  = """ & lgPageNo   & """"			& vbCr 
    Response.Write "	.frm1.hdnMvmtFrDt.value		= """ & Trim(ConvSPChars(Request("txtMvmtFrDt")))           & """" & vbCr
    Response.Write "	.frm1.hdnMvmtToDt.value		= """ & Trim(ConvSPChars(Request("txtMvmtToDt")))           & """" & vbCr
    Response.Write "	.frm1.hdnIvFrDt.value		= """ & Trim(ConvSPChars(Request("txtIvFrDt")))             & """" & vbCr
    Response.Write "	.frm1.hdnIvToDt.value		= """ & Trim(ConvSPChars(Request("txtIvToDt")))             & """" & vbCr
    Response.Write "	.frm1.hdnMvmtType.value		= """ & Trim(ConvSPChars(Request("txtMvmtType")))           & """" & vbCr
    Response.Write "	.frm1.hdnrdoflg.value		= """ & Trim(ConvSPChars(Request("rdoAppflg")))				& """" & vbCr
	
	Response.Write "	.frm1.txtMvmtTypeNm.value	= """ & ConvSPChars(iStrIoTypeNm)              				& """" & vbCr
	
	Response.Write "	.frm1.txtTotMvmtAmt.value	= """ & ConvSPChars(iSumMvmtAmt) & """" & vbCr
	Response.Write "	.frm1.txtTotIvAmt.value		= """ & ConvSPChars(iSumGrirAmt) & """" & vbCr
	Response.Write "	.frm1.txtTotBalanceAmt.value= """ & ConvSPChars(iSumMvmtAmt) - ConvSPChars(iSumGrirAmt) & """" & vbCr
	Response.Write "    .DbQueryOk"										& vbCr
	Response.Write "    .frm1.vspdData.Redraw = True   "                & vbCr   	
	Response.Write "End With"											& vbCr
    Response.Write "</Script>"											& vbCr        

End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Dim iLoopCount  
	Dim TmpAmt
	
	Const C_SHEETMAXROWS_D  = 100
     iPrevEndRow = 0
    If CLng(lgPageNo) > 0 Then
		iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

   iLoopCount = -1
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
  
   Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
		
		TmpAmt = Cdbl(rs0("mvmt_loc_amt")) - cdbl(rs0("grir_amt"))
		
        iRowStr  = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("mvmt_rcpt_no"))	                       
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("mvmt_dt"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("io_type_cd"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("io_type_nm"))
		iRowStr = iRowStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("mvmt_loc_amt"),0)
		iRowStr = iRowStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(rs0("grir_amt"),0)
		iRowStr = iRowStr & Chr(11) & UniConvNumDBToCompanyWithOutChange(TmpAmt,0)
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("mvmt_biz_area"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("biz_area_nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("bp_cd"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("bp_nm"))
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("gm_no"))
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount - 1

        If iLoopCount < C_SHEETMAXROWS_D Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = istrData	
		   istrData = ""
		   iEndRow = iPrevEndRow + iLoopCount + 1
        Else
           lgPageNo = lgPageNo + 1
           iEndRow = iPrevEndRow + iLoopCount
           Exit Do
        End If
	    rs0.MoveNext
	Loop
	
	
	iLngRow = iLoopCount
	iTotstrData = Join(PvArr, "")
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

%>
