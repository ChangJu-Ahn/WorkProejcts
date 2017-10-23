<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3521qb1
'*  4. Program Name         : 미발주구매요청현황조회 
'*  5. Program Desc         : 미발주구매요청현황조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/06/27
'*  8. Modified date(Last)  : 2003/06/27
'*  9. Modifier (First)     : KANG SU HWAN
'* 10. Modifier (Last)      : KANG SU HWAN
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
	Dim lgPageNo
	Dim lgDataExist
	Dim lgPlantNm,lgItemNm,lgReqDeptNm	
	
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")
     
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
    Const C_SHEETMAXROWS_D = 100 

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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

	lgstrData  = Join(PvArr, "")

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
	Dim lgStrSql
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,3)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    '---공장 
    If Len(Trim(Request("txtPlantCd"))) Then
    	lgStrSql = lgStrSql & " AND A.PLANT_CD =  " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
    End If
    
    '---품목 
    If Len(Trim(Request("txtItemCd"))) Then
    	lgStrSql = lgStrSql & " AND A.ITEM_CD =  " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
    End If

    '---요청부서 
    If Len(Trim(Request("txtRqDeptCd"))) Then
    	lgStrSql = lgStrSql & " AND A.REQ_DEPT =  " & FilterVar(Trim(UCase(Request("txtRqDeptCd"))), " " , "S") & " "
    End If

    '---구매요청일 
    If Len(Trim(Request("txtPrFrDt"))) Then
    	lgStrSql = lgStrSql & " AND A.REQ_DT >=  " & FilterVar(uniConvDate(Trim(Request("txtPrFrDt"))), "''", "S") & ""
    Else
    	lgStrSql = lgStrSql & " AND A.REQ_DT >= " & FilterVar("1900-01-01", "''", "S") & ""
    End If

    If Len(Trim(Request("txtPrToDt"))) Then
    	lgStrSql = lgStrSql & " AND A.REQ_DT <=  " & FilterVar(uniConvDate(Trim(Request("txtPrToDt"))), "''", "S") & ""
    Else
    	lgStrSql = lgStrSql & " AND A.REQ_DT <= " & FilterVar("2999-12-31", "''", "S") & ""
    End If

    '---필요납기일 
    If Len(Trim(Request("txtPdFrDt"))) Then
    	lgStrSql = lgStrSql & " AND A.DLVY_DT >=  " & FilterVar(uniConvDate(Request("txtPdFrDt")), "''", "S") & ""
    Else
    	lgStrSql = lgStrSql & " AND A.DLVY_DT >= " & FilterVar("1900-01-01", "''", "S") & ""
    End If

    If Len(Trim(Request("txtPdToDt"))) Then
    	lgStrSql = lgStrSql & " AND A.DLVY_DT <=  " & FilterVar(uniConvDate(Request("txtPdToDt")), "''", "S") & ""
    Else
    	lgStrSql = lgStrSql & " AND A.DLVY_DT <= " & FilterVar("2999-12-31", "''", "S") & ""
    End If

     UNISqlId(0) = "M3521QA101"
     UNISqlId(1) = "M2111QA302"								              '공장명 
	 UNISqlId(2) = "M2111QA303"											  '품목명 
	 UNISqlId(3) = "M2111QA305"											  '부서명 

     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
     UNIValue(0,1)  = " " & FilterVar(Trim(UCase(Request("txtchangorgid"))), " " , "S") & " "
     UNIValue(0,2)  = lgStrSql
     UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Order By 조건 
	 UNIValue(1,0) = " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
     UNIValue(2,0) = " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
     UNIValue(2,1) = " " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
     UNIValue(3,0) = " " & FilterVar(Trim(UCase(Request("txtRqDeptCd"))), " " , "S") & " "
     UNIValue(3,1) = " " & FilterVar(Trim(UCase(Request("txtchangorgid"))), " " , "S") & " "

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

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
		lgPlantNm = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
			Call DisplayMsgBox("122700", vbInformation, "X", "X", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			FalsechkFlg = True	
	       	rs0.Close
	       	Set rs0 = Nothing
			Exit Sub		'20030124 - leejt
		End If
    Else    
		lgItemNm = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtRqDeptCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "요청부서", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		lgReqDeptNm = rs3(1)
        rs3.Close
        Set rs3 = Nothing
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
         .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
  		 
  		 .frm1.txtHPlantCd.value   = "<%=ConvSPChars(Request("txtPlantCd"))%>"
         .frm1.txtHItemCd.value    = "<%=ConvSPChars(Request("txtItemCd"))%>"
  		 .frm1.txtHRqDeptCd.value   = "<%=ConvSPChars(Request("txtRqDeptCd"))%>"
         .frm1.txtHPrFrDt.value     = "<%=ConvSPChars(Request("txtPrFrDt"))%>"
         .frm1.txtHPrToDt.value     = "<%=ConvSPChars(Request("txtPrToDt"))%>"
         .frm1.txtHPdFrDt.value     = "<%=ConvSPChars(Request("txtPdFrDt"))%>"
         .frm1.txtHPdToDt.value     = "<%=ConvSPChars(Request("txtPdToDt"))%>"

         .frm1.txtPlantNm.value		=  "<%=ConvSPChars(lgPlantNm)%>" 	
  		 .frm1.txtItemNm.value		=  "<%=ConvSPChars(lgItemNm)%>" 	
  		 .frm1.txtRqDeptNm.value	=  "<%=ConvSPChars(lgReqDeptNm)%>" 	
  		 
         .DbQueryOk
	End with
</Script>	
<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>

