<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5211qb2
'*  4. Program Name         : 선적상세조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jin-hyun Shin
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : 
'=======================================================================================================
Option Explicit
%>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next

Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgStrPrevKey                                            '☜ : 이전 값 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim iTotstrData

Dim ICount  		                                        '   Count for column index
Dim strIncotermsCd												'가격조건 
Dim strIncotermsCdFrom				
Dim strIncotermsLookUp
Dim strPurGrpCd												'	구매그룹 
Dim strPurGrpCdFrom 										
Dim strBpCd													'   수출자 
Dim strBpCdFrom
Dim strBlFrDt                                               '   Bl접수일 
Dim strBlToDt
Dim strLoadingFrDt                                          '   선적일 
Dim strLoadingToDt
Dim strCfmFlg								                '   확정여부 
Dim strCfmFlgFrom	
Dim strItemCd								                '   품목 
Dim strItemCdFrom	
Dim strPlantCd								                '   공장 
Dim strPlantCdFrom	
Dim strBlNo
Dim StrBlNoFrom
Dim lgPageNo
Dim lgDataExist

Dim arrRsVal(11)											'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array
Dim iFrPoint
iFrPoint=0

Const Major_Cd_Incoterms = "B9006"


    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

	lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 

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
       iFrPoint     = C_SHEETMAXROWS_D * CLng(lgPageNo) 
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()
	Dim iStrSql
    Redim UNISqlId(7)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(6,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
    Dim strIncotermsCd
    
	iStrSql=""
    '---품목 
    If Len(Trim(Request("txtItemCd"))) Then
		iStrSql = iStrSql & " AND A.ITEM_CD =  " & FilterVar(Request("txtItemCd"), "''", "S") & ""
    End If
     '---공장 
    If Len(Trim(Request("txtPlantCd"))) Then
		iStrSql = iStrSql & " AND A.PLANT_CD =  " & FilterVar(Request("txtPlantCd"), "''", "S") & ""
    End If
     '---가격조건 
    If Len(Trim(Request("txtIncotermsCd"))) Then
		iStrSql = iStrSql & " AND B.INCOTERMS =  " & FilterVar(Request("txtIncotermsCd"), "''", "S") & ""
    End If
     '---구매그룹 
    If Len(Trim(Request("txtPurGrpCd"))) Then
		iStrSql = iStrSql & " AND B.PUR_GRP =  " & FilterVar(Request("txtPurGrpCd"), "''", "S") & ""
    End If
     '---수출자 
    If Len(Trim(Request("txtBpCd"))) Then
		iStrSql = iStrSql & " AND B.BENEFICIARY =  " & FilterVar(Request("txtBpCd"), "''", "S") & ""
    End If
     '---bl접수일 
    If Len(Trim(Request("txtBlFrDt"))) Then
		iStrSql = iStrSql & " AND B.BL_ISSUE_DT >=  " & FilterVar(UNIConvDate(Trim(Request("txtBlFrDt"))), "''", "S") & ""
    Else
		iStrSql = iStrSql & " AND B.BL_ISSUE_DT >=  " & FilterVar(UNIConvDate("1900-01-01"), "''", "S") & ""
    End If
    If Len(Trim(Request("txtBlToDt"))) Then
		iStrSql = iStrSql & " AND B.BL_ISSUE_DT <=  " & FilterVar(UNIConvDate(Trim(Request("txtBlToDt"))), "''", "S") & ""
    Else
		iStrSql = iStrSql & " AND B.BL_ISSUE_DT <=  " & FilterVar(UNIConvDate("2999-12-30"), "''", "S") & ""
    End If  
     '---선적일 
    If Len(Trim(Request("txtLoadingFrDt"))) Then
		iStrSql = iStrSql & " AND B.LOADING_DT >=  " & FilterVar(UNIConvDate(Trim(Request("txtLoadingFrDt"))), "''", "S") & ""
    Else
		iStrSql = iStrSql & " AND B.LOADING_DT >=  " & FilterVar(UNIConvDate("1900-01-01"), "''", "S") & ""
    End If
	
    If Len(Trim(Request("txtLoadingToDt"))) Then
		iStrSql = iStrSql & " AND B.LOADING_DT <=  " & FilterVar(UNIConvDate(Trim(Request("txtLoadingToDt"))), "''", "S") & ""
    Else
		iStrSql = iStrSql & " AND B.LOADING_DT <=  " & FilterVar(UNIConvDate("2999-12-30"), "''", "S") & ""
    End If       
     '---확정여부 
    If Len(Trim(Request("txtCfmFlg"))) Then
		iStrSql = iStrSql & " AND B.POSTING_FLG =  " & FilterVar(Request("txtCfmFlg"), "''", "S") & ""
    End If
     '---B/L No
    If Len(Trim(Request("txtBlNo"))) Then
		iStrSql = iStrSql & " AND A.BL_NO =  " & FilterVar(Request("txtBlNo"), "''", "S") & ""
    End If

     If Request("gBizArea") <> "" Then
        iStrSql = iStrSql & " AND B.BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        iStrSql = iStrSql & " AND B.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        iStrSql = iStrSql & " AND B.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If

    
    If Len(Trim(Request("txtIncotermsCd"))) Then
    	strIncotermsCd	= FilterVar(Trim(UCase(Request("txtIncotermsCd"))), " " , "S")
    Else
    	strIncotermsCd	= FilterVar("zzzzzzzzz", " " , "S")
    End If

	 UNISqlId(0) = "M4214QA201"
     UNISqlId(1) = "M3111QA102"								              '공급처명 
	 UNISqlId(2) = "M2111QA302"											  '공장 
     UNISqlId(3) = "M2111QA303"								              '품목 
     UNISqlId(4) = "M3111QA104"								              '구매그룹명 
	 UNISqlId(5) = "S0000QA000"											  '가격조건 
	 UNISqlId(6) = "M4214QA202"											  '촐수량 

     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
     UNIValue(0,1)  = iStrSql
     UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By 조건 
     UNIValue(1,0)  = " " & FilterVar(Request("txtBpCd"), "''", "S") & ""
     UNIValue(2,0)  = " " & FilterVar(Request("txtPlantCd"), "''", "S") & ""
     UNIValue(3,0)  = " " & FilterVar(Request("txtPlantCd"), "''", "S") & ""
     UNIValue(3,1)  = " " & FilterVar(Request("txtItemCd"), "''", "S") & ""
     UNIValue(4,0)  = " " & FilterVar(Request("txtPurGrpCd"), "''", "S") & ""
	 UNIValue(5,0)  = " " & FilterVar(UCase(Trim(Major_Cd_Incoterms)), "''", "S") & ""
     UNIValue(5,1)  = strIncotermsCd
     
     UNIValue(6,0)  = iStrSql
     
     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)			
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    Dim FalsechkFlg
    
    FalsechkFlg = False 

    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수출자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs1(0)
		arrRsVal(5) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
	
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs2(0)
		arrRsVal(1) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
	
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "X", "X", I_MKSCRIPT)
	       Set rs0 = Nothing
	       Exit Sub
'		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
'	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(6) = rs3(0)
		arrRsVal(7) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
        If Len(Request("txtPurGrpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs4(0)
		arrRsVal(3) = rs4(1)
        rs4.Close
        Set rs2 = Nothing
    End If
    
    If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtIncotermsCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "가격조건", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If

    If  rs6.EOF And rs6.BOF Then
        rs6.Close
        Set rs6 = Nothing
    Else    
		arrRsVal(10) = rs6(0)
        rs6.Close
        Set rs6 = Nothing
    End If

	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
    

%>

<Script Language=vbscript>
    
    With Parent
        .ggoSpread.Source  = .frm1.vspdData
        Parent.frm1.vspdData.Redraw = False
        .ggoSpread.SSShowData "<%=iTotstrData%>","F"                  '☜ : Display data
        
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",9),.GetKeyPos("A",7),"C","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",9),.GetKeyPos("A",8),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",10),"D","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.parent.gCurrency,.GetKeyPos("A",11),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",9),.GetKeyPos("A",20),"D","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",9),.GetKeyPos("A",21),"D","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",9),.GetKeyPos("A",23),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",9),.GetKeyPos("A",24),"A","Q","X","X")
        
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag

		.frm1.hdnBeneficiaryCd.value		= "<%=ConvSPChars(Request("txtBpCd"))%>"
        .frm1.hdnIncotermsCd.value	= "<%=ConvSPChars(Request("txtIncotermsCd"))%>"
        .frm1.hdnPurGrpCd.value		= "<%=ConvSPChars(Request("txtPurGrpCd"))%>"
		.frm1.hdnBlIssueFrDt.value	= "<%=Request("txtBlFrDt")%>"
        .frm1.hdnBlIssueToDt.value	= "<%=Request("txtBlToDt")%>"
        .frm1.hdnLoadingFrDt.value	= "<%=ConvSPChars(Request("txtLoadingFrDt"))%>"
        .frm1.hdnLoadingToDt.value	    = "<%=ConvSPChars(Request("txtLoadingToDt"))%>"
		.frm1.hdnstrCfmFlg.value	    = "<%=ConvSPChars(Request("txtCfmFlg"))%>"
		.frm1.hdnItemCd.value	= "<%=ConvSPChars(Request("txtItemCd"))%>"
        .frm1.hdnPlantCd.value	    = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hdnBlNo.value	    = "<%=ConvSPChars(Request("txtBlNo"))%>"
		
		.frm1.txtBeneficiaryNm.value	=  "<%=ConvSPChars(arrRsVal(5))%>" 	
		.frm1.txtPlantNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>"
		.frm1.txtitemNm.value			=  "<%=ConvSPChars(arrRsVal(7))%>"
		.frm1.txtPurGrpNm.value			=  "<%=ConvSPChars(arrRsVal(3))%>"
		.frm1.txtIncotermsNm.value		=  "<%=ConvSPChars(arrRsVal(9))%>"
		.frm1.txtTotQty.Text			=  "<%=UNINumClientFormat(arrRsVal(10), ggQty.DecPoint, 0)%>" 

		.DbQueryOk
		Parent.frm1.vspdData.Redraw = True
	End with

</Script>	

<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>
