<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2002/01/18
'*  7. Modified date(Last)  : 2003/05/20
'*  8. Modifier (First)     : park jin uk
'*  9. Modifier (Last)      : Lee Eun Hee
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
<%
On Error Resume Next

Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4, rs5     							'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgStrPrevKey                                            '☜ : 이전 값 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim iTotstrData

Dim ICount  		                                        '   Count for column index
Dim strPlantCd												'	공장 
Dim strPlantCdFrom				
Dim strPurGrpCd				                  				'	구매그룹 
Dim strPurGrpCdFrom 										
Dim strBeneficiary                                          '   수출자 
Dim strBeneficiaryFrom
Dim strFrDt                                                 '   개설일 
Dim strToDt
Dim strPayMeth                                              '   결재방법 
Dim strPayMethFrom
Dim strIncoterms                                            '   가격조건 
Dim strIncotermsFrom
Dim arrRsVal(12)											'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array	
Dim lgPageNo
Dim lgDataExist
Dim iFrPoint
iFrPoint=0


    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

	 lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 

     Call  TrimData()                                                     '☜ : Parent로 부터의 데이타 가공 
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

    Redim UNISqlId(6)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(6,14)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "M3211QA101"

     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
     
     UNISqlId(1) = "M2111QA302"								              '공장명 
     UNISqlId(2) = "M3111QA104"								              '구매그룹명 
     UNISqlId(3) = "M3211QA102"								              '수출자명     
     UNISqlId(4) = "M3211QA103"											  '결재방법명        
     UNISqlId(5) = "M3211QA104"								              '가격조건명     	
																		  'Reusage is Recommended

	 UNIValue(0,1)  = UCase(Trim(strPlantCdFrom))		'---공장 
	 UNIValue(0,2)  = UCase(Trim(strPlantCd))
	 UNIValue(0,3)  = UCase(Trim(strPurGrpCdFrom))      '---구매그룹 
     UNIValue(0,4)  = UCase(Trim(strPurGrpCd))
     UNIValue(0,5)  = UCase(Trim(strBeneficiaryFrom))	'---수출자 
     UNIValue(0,6)  = UCase(Trim(strBeneficiary))     
     UNIValue(0,7)  = " " & FilterVar(UNIConvDate(strFrDt), "''", "S") & ""	'---개설일 
     UNIValue(0,8)  = " " & FilterVar(UNIConvDate(strToDt), "''", "S") & ""   
     UNIValue(0,9)  = UCase(Trim(strPayMethFrom))	   	'---결재방법 
     UNIValue(0,10) = UCase(Trim(strPayMeth))
     UNIValue(0,11) = UCase(Trim(strIncotermsFrom))     '---가격조건 
     UNIValue(0,12) = UCase(Trim(strIncoterms))
     
     UNIValue(1,0)  = UCase(Trim(strPlantCd))
     UNIValue(2,0)  = UCase(Trim(strPurGrpCd))  
     UNIValue(3,0)  = UCase(Trim(strBeneficiary))           
     UNIValue(4,0)  = UCase(Trim(strPayMeth))
     UNIValue(5,0)  = UCase(Trim(strIncoterms))

     
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
    '============================= 추가된 부분 =====================================================================
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If

    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
         If Len(Request("txtPurGrpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtBeneficiary")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수출자", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
        If Len(Request("txtPayMeth")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "결재방법", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If
    
    If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtIncoterms")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "가격조건", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If
    
	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()


     '---공장 
    If Len(Trim(Request("txtPlantCd"))) Then
    	strPlantCd	= " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
    	strPlantCdFrom = strPlantCd
    Else
    	strPlantCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPlantCdFrom = "''"
    End If
     '---구매그룹 
    If Len(Trim(Request("txtPurGrpCd"))) Then
    	strPurGrpCd	= " " & FilterVar(Trim(UCase(Request("txtPurGrpCd"))), " " , "S") & " "
    	strPurGrpCdFrom = strPurGrpCd
    Else
    	strPurGrpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPurGrpCdFrom = "''"
    End If
     '---수출자 
    If Len(Trim(Request("txtBeneficiary"))) Then
    	strBeneficiary	= " " & FilterVar(Trim(UCase(Request("txtBeneficiary"))), " " , "S") & " "
    	strBeneficiaryFrom = strBeneficiary
    Else
    	strBeneficiary	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strBeneficiaryFrom = "''"    	
    End If
     '---개설일 
    If Len(Trim(Request("txtFrDt"))) Then
    	strFrDt 	= "" & Trim(Request("txtFrDt")) & ""
    Else
    	strFrDt	= UNIDateClientFormat("1900-01-01")
    End If

    If Len(Trim(Request("txtToDt"))) Then
    	strToDt 	= "" & Trim(Request("txtToDt")) & ""
    Else
    	strToDt	= UNIDateClientFormat("2999-12-30")
    End If    
    '---결재방법 
    If Len(Trim(Request("txtPayMeth"))) Then
    	strPayMeth	= " " & FilterVar(Trim(UCase(Request("txtPayMeth"))), " " , "S") & " "
    	strPayMethFrom = strPayMeth
    Else
    	strPayMeth	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strPayMethFrom = "''"
    End If

    '---가격조건 
    If Len(Trim(Request("txtIncoterms"))) Then
    	strIncoterms	= " " & FilterVar(Trim(UCase(Request("txtIncoterms"))), " " , "S") & " "
    	strIncotermsFrom = strIncoterms
    Else
    	strIncoterms	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
    	strIncotermsFrom = "''"
    End If

End Sub

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         Parent.frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>","F"                  '☜ : Display data

		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",10),.GetKeyPos("A",11),"A","Q","X","X")
         
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
        
		.frm1.hdnPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
        .frm1.hdnPurGrpCd.value		= "<%=ConvSPChars(Request("txtPurGrpCd"))%>"
        .frm1.hdnBeneficiary.value	= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
        .frm1.hdnFrDt.value			= "<%=Request("txtFrDt")%>"
        .frm1.hdnToDt.value			= "<%=Request("txtToDt")%>"
        .frm1.hdnPayMeth.value		= "<%=ConvSPChars(Request("txtPayMeth"))%>"
        .frm1.hdnIncoterms.value	= "<%=ConvSPChars(Request("txtIncoterms"))%>"
        		 
		 .frm1.txtPlantNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		 .frm1.txtPurGrpNm.value	=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		 .frm1.txtBpNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>" 	
  		 .frm1.txtPayMethNm.value	=  "<%=ConvSPChars(arrRsVal(7))%>" 	
  		 .frm1.txtIncotermsNm.value	=  "<%=ConvSPChars(arrRsVal(9))%>"
  		 
         .DbQueryOk
         Parent.frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>
