<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MB112QA1
'*  4. Program Name         : 사급품미출고조회 
'*  5. Program Desc         : 사급품미출고조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/06/30
'*  8. Modified date(Last)  : 2003/06/30
'*  9. Modifier (First)     : Kang Su Hwan
'* 10. Modifier (Last)      : Kang Su Hwan
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
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	On Error Resume Next

	Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
	Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
	Dim rs1, rs2, rs3, rs4, rs5, rs6										   '☜ : DBAgent Parameter 선언 
	Dim lgstrData                                                              '☜ : data for spreadsheet data
	Dim iTotstrData        
	
	Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgStrSql	
	Dim lgDataExist
	Dim lgPageNo
	
	Dim arrRsVal(3)														   '* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array	
	
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100

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

    Dim strVal
    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(5,2)

    UNISqlId(0) = "MB112QA101"
    UNISqlId(1) = "M2111QA302"								              '공장명 
    UNISqlId(2) = "M3111QA102"								              '거래처명 
    UNISqlId(3) = "M4111QA502"											  '창고명 
    UNISqlId(4) = "M2111QA303"											  '품목명   
	
	lgStrSql = lgStrSql & " AND A.PLANT_CD =  " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	If Len(Request("txtSpplCd")) Then
		lgStrSql = lgStrSql & " AND A.SPPL_CD =  " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "
	End If
    If Len(Request("txtSlCd")) Then
		lgStrSql = lgStrSql & " AND A.SL_CD =  " & FilterVar(UCase(Request("txtSlCd")), "''", "S") & " "
	End If		
    If Len(Request("txtItemCd")) Then
		lgStrSql = lgStrSql & " AND A.ITEM_CD =  " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
	End If	
    
	UNIValue(0,0) = lgSelectList                                          '☜: Select list
	UNIValue(0,1) = lgStrSql                         
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNIValue(1,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(2,0)  = " " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "
    UNIValue(3,0)  = " " & FilterVar(UCase(Request("txtSlCd")), "''", "S") & " "      
    UNIValue(4,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(4,1)  = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4)
    
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
        If Len(Request("txtSpplCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtSlCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "창고", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
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
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "X", "X", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(1) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
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
End Sub


%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>"                            '☜: Display data 
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
         
         .frm1.hdnPlantCd.value    = "<%=ConvSPChars(Request("txtPlantCd"))%>"
         .frm1.hdnItemCd.value     = "<%=ConvSPChars(Request("txtItemCd"))%>"
         .frm1.hdnSlCd.value       = "<%=ConvSPChars(Request("txtSlCd"))%>"
         .frm1.hdnSpplCd.value     = "<%=ConvSPChars(Request("txtSpplCd"))%>"
         
         .frm1.txtPlantNm.value		=  "<%=ConvSPChars(arrRsVal(0))%>" 	
  		 .frm1.txtItemNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		 .frm1.txtSlNm.value		=  "<%=ConvSPChars(arrRsVal(2))%>" 	
  		 .frm1.txtSpplNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
         .DbQueryOk
         .frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>

