<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         : 구매입출고관리- 출고상세조회 
'*  6. Modified date(First) : 2000/12/12
'*  7. Modified date(Last)  : 2003/06/02
'*  8. Modifier (First)     : ByunJiHyun
'*  9. Modifier (Last)      : Kim Jin Ha
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
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
Dim iTotstrData
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

Dim lgDataExist
Dim lgPageNo

Dim strPlantCd												'	공장 
Dim strPlantCdFrom	
Dim strItemCd                                               '   품목 
Dim strItemCdFrom
Dim strBpCd												    '   거래처 
Dim strBpCdFrom		
Dim strMvFrDt                                               '   출고일 
Dim strMvToDt
Dim strSlCd													'	창고 
Dim strSlCdFrom 										
Dim strIoType                                               '   출고유형 
Dim strIoTypeFrom	
Dim strPoNo													'	발주번호 
'Dim strPoNoFrom

Dim arrRsVal(11)											'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array	

Dim iPrevEndRow
Dim iEndRow

	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
	iPrevEndRow = 0
	iEndRow = 0	 
	
    Call  TrimData()                                                     '☜ : Parent로 부터의 데이타 가공 
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()

    Redim UNISqlId(6)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(6,14)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "M4111QA801"
     
     UNISqlId(1) = "M2111QA302"								              '공장명 
     UNISqlId(2) = "M2111QA303"											  '품목명     
     UNISqlId(3) = "M3111QA102"								              '거래처명 
	 UNISqlId(4) = "M4111QA502"											  '창고명 
	 UNISqlId(5) = "M4111QA702"											  '출고유형명 
																		  'Reusage is Recommended
     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	 UNIValue(0,1)  = UCase(Trim(strPlantCdFrom))		'---공장 
	 UNIValue(0,2)  = UCase(Trim(strPlantCd))
	 UNIValue(0,3)  = UCase(Trim(strItemCdFrom))		'---품목 
     UNIValue(0,4)  = UCase(Trim(strItemCd))
     UNIValue(0,5)  = UCase(Trim(strBpCdFrom))			'---거래처 
     UNIValue(0,6)  = UCase(Trim(strBpCd))  
     UNIValue(0,7)  = UCase(Trim(strMvFrDt))			'---출고일 
     UNIValue(0,8)  = UCase(Trim(strMvToDt)) 
	 UNIValue(0,9)  = UCase(Trim(strSlCdFrom))		    '---창고 
     UNIValue(0,10)  = UCase(Trim(strSlCd))     
     UNIValue(0,11)  = UCase(Trim(strIoTypeFrom))		'---출고형태    
     UNIValue(0,12)  = UCase(Trim(strIoType))
     UNIValue(0,13)  = UCase(Trim(strPoNo))				'---발주번호    
     'UNIValue(0,14)  = UCase(Trim(strPoNo))
     
     UNIValue(1,0)  = UCase(Trim(strPlantCd))
     UNIValue(2,0)  = UCase(Trim(strPlantCd))
     UNIValue(2,1)  = UCase(Trim(strItemCd))  
     UNIValue(3,0)  = UCase(Trim(strBpCd))      
     UNIValue(4,0)  = UCase(Trim(strSlCd))
     UNIValue(5,0)  = UCase(Trim(strIoType))
     
     UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By 조건 
	 UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
		 
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    Dim FalsechkFlg
    
    FalsechkFlg = False 
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
	       Set rs2 = Nothing
	       FalsechkFlg = True	
	       Exit Sub
		End If
    Else    
		arrRsVal(1) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "거래처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
        If Len(Request("txtSlCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "창고", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(3) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If
    
    If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtIoType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "출고유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        'Response.Write "<Script Language=vbscript>" & vbCr
		'Response.write "parent.frm1.vspdData.MaxRows = 0 " & chr(13)
		'Response.Write "</Script>"
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     '---공장 
    If Len(Trim(Request("txtPlantCd"))) Then
    	strPlantCd	= " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    	strPlantCdFrom = strPlantCd
    Else
    	strPlantCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPlantCdFrom = "''"
    End If
    '---품목 
    If Len(Trim(Request("txtItemCd"))) Then
    	strItemCd	= " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
    	strItemCdFrom = strItemCd
    Else
    	strItemCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strItemCdFrom = "''"
    End If
    '---거래처 
    If Len(Trim(Request("txtBpCd"))) Then
    	strBpCd	= " " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
    	strBpCdFrom = strBpCd
    Else
    	strBpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strBpCdFrom = "''"    	
    End If
     '---출고일 
    If Len(Trim(Request("txtMvFrDt"))) Then
    	strMvFrDt 	= " " & FilterVar(uniConvDate(Request("txtMvFrDt")), "''", "S") & ""
    Else
    	strMvFrDt	= "" & FilterVar("1900/01/01", "''", "S") & ""
    End If

    If Len(Trim(Request("txtMvToDt"))) Then
    	strMvToDt 	= " " & FilterVar(uniConvDate(Request("txtMvToDt")), "''", "S") & ""
    Else
    	strMvToDt	= "" & FilterVar("2999/12/30", "''", "S") & ""
    End If    
     '---창고 
    If Len(Trim(Request("txtSlCd"))) Then
    	strSlCd	= " " & FilterVar(UCase(Request("txtSlCd")), "''", "S") & " "
    	strSlCdFrom = strSlCd
    Else
    	strSlCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strSlCdFrom = "''"
    End If
     '---출고유형 
    If Len(Trim(Request("txtIoType"))) Then
    	strIoType	= " " & FilterVar(UCase(Request("txtIoType")), "''", "S") & " "
    	strIoTypeFrom = strIoType
    Else
    	strIoType	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strIoTypeFrom = "''"
    End If
     '---발주번호 
    If Len(Trim(Request("txtPoNo"))) Then
    	strPoNo	= " AND A.PO_NO = " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & " "
    	'strPoNoFrom = strPoNo
    Else
    	strPoNo	= ""
    	'strPoNoFrom = "''"
    End If
End Sub

     
%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         .frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>","F"                  '☜ : Display data
         
         Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",16),.GetKeyPos("A",17),"A","Q","X","X")
         Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.parent.gCurrency,.GetKeyPos("A",18),"A","Q","X","X")
         
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
         .frm1.hdnPlantCd.value    = "<%=ConvSPChars(Request("txtPlantCd"))%>"
         .frm1.hdnItemCd.value     = "<%=ConvSPChars(Request("txtItemCd"))%>"
         .frm1.hdnBpCd.value       = "<%=ConvSPChars(Request("txtBpCd"))%>"
         .frm1.hdnMvFrDt.value     = "<%=Request("txtMvFrDt")%>"
         .frm1.hdnMvToDt.value     = "<%=Request("txtMvToDt")%>"
         .frm1.hdnSlCd.value       = "<%=ConvSPChars(Request("txtSlCd"))%>"
         .frm1.hdnIoType.value     = "<%=ConvSPChars(Request("txtIoType"))%>"
         .frm1.hdnPoNo.value       = "<%=ConvSPChars(Request("txtPoNo"))%>"
         
         .frm1.txtPlantNm.value			=  "<%=ConvSPChars(arrRsVal(0))%>" 	
  		 .frm1.txtItemNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		 .frm1.txtBpNm.value			=  "<%=ConvSPChars(arrRsVal(2))%>" 	
  		 .frm1.txtSlNm.value			=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		 .frm1.txtIoTypeNm.value		=  "<%=ConvSPChars(arrRsVal(4))%>"
         .DbQueryOk
         .frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>
