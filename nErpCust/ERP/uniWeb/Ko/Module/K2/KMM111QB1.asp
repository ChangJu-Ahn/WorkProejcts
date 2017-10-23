<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MM111QA1
'*  4. Program Name         : 멀티컴퍼니매입조회 
'*  5. Program Desc         : 멀티컴퍼니매입조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/07/02
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")


'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrPoNo
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim iLngMaxRow		' 현재 그리드의 최대Row
Dim iLngRow
Dim GroupCount  
Dim lgCurrency        
Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수     
Dim lgDataExist
Dim lgPageNo
Dim SoCompanyNm			'☜ : 수주법인 
Dim ItemNm			'☜ : 품목명 
Dim PrTypeNm		'☜ : 요청구분명 
Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows
Dim iStrPostingFlg

intARows=0
intTRows=0
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status


Call HideStatusWnd                                                               '☜: Hide Processing message
lgOpModeCRUD  = Request("txtMode") 

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
		 Call  SubBizQueryMulti()
End Select

Sub SubBizQueryMulti()


	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

	Call FixUNISQLData()		'☜ : DB-Agent로 보낼 parameter 데이타 set
	Call QueryData()			'☜ : DB-Agent를 통한 ADO query
	
	'-----------------------
	'Result data display area
	'----------------------- 

%>

	<Script Language=vbscript>
		With parent
			.frm1.txtSoCompanyCd.value = "<%=ConvSPChars(Request("txtSoCompanyCd"))%>"			
			.frm1.txtSoCompanyNm.Value	= "<%=SoCompanyNm%>"							
			.frm1.txtSoCompanyCd.focus
			
			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then
				
				'Show multi spreadsheet data from this line
				       
				.ggoSpread.Source    = .frm1.vspdData 
				.ggoSpread.SSShowData "<%=istrData%>"                  '☜: Display data 
				
				.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
				
				.DbQueryOk <%=intARows%>,<%=intTRows%>
							
			End If
		End with
	</Script>	
<%	
End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim PvArr
	Const C_SHEETMAXROWS_D  = 100            
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt

	Const	M_IV_HDR_BP_CD					= 0
	Const	M_IV_HDR_BP_NM                  = 1
	Const	M_IV_HDR_SPPL_IV_NO             = 2
	Const	M_IV_HDR_IV_NO                  = 3
	Const	M_IV_HDR_POSTED_FLG             = 4
	Const	M_IV_HDR_IV_DT                  = 5
	Const	M_IV_HDR_PUR_GRP                = 6
	Const	M_IV_HDR_PUR_GRP_NM             = 7
	Const	M_IV_HDR_IV_CUR                 = 8
	Const	M_IV_HDR_NET_DOC_AMT            = 9
	Const	M_IV_HDR_TOT_VAT_DOC_AMT        = 10
	Const	M_IV_HDR_GROSS_DOC_AMT          = 11
	Const	M_IV_HDR_VAT_TYPE               = 12
	Const	M_IV_HDR_VAT_TYPE_NM            = 13
	Const	M_IV_HDR_VAT_RT                 = 14
	Const	M_IV_HDR_PAY_METH               = 15
	Const	M_IV_HDR_PAY_METH_NM            = 16
	Const	M_IV_HDR_PAY_MENT_TERM          = 17
	Const	M_IV_HDR_PAY_MENT_TERM_NM	    = 18   
	Const	M_IV_HDR_GL_NO	    = 19

    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
	
	'----- 레코드셋 칼럼 순서 ----------
	'A.BP_CD, (SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD = A.BP_CD) BP_NM, A.SPPL_IV_NO,
	'A.IV_NO, A.POSTED_FLG, A.IV_DT, A.PUR_GRP, (SELECT PUR_GRP_NM FROM B_PUR_GRP WHERE PUR_GRP = A.PUR_GRP) PUR_GRP_NM,
 	'A.IV_CUR, A.NET_DOC_AMT, A.TOT_VAT_DOC_AMT, A.GROSS_DOC_AMT, A.VAT_TYPE, 
	'(SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'B9001' AND MINOR_CD = A.VAT_TYPE) VAT_TYPE_NM,
	'A.VAT_RT, A.PAY_METH, (SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'B9004' AND MINOR_CD = A.PAY_METH) PAY_METH_NM,
	'A.PAYMENT_TERM, (SELECT MINOR_NM FROM B_MINOR WHERE MAJOR_CD = 'B9006' AND MINOR_CD = A.PAYMENT_TERM) PAYMENT_TERM_NM
	'-----------------------------------
	
	iLoopCount = 0
    ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_BP_CD))												'수주법인              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_BP_NM))											    '수주법인명            
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_SPPL_IV_NO))										    '공급처세금계산서번호  
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_IV_NO))											    '매입번호              
		IF rs0(M_IV_HDR_POSTED_FLG) = "Y" Then
			iStrPostingFlg = "확정"
		Else
			iStrPostingFlg = "미확정"			
		End if		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(iStrPostingFlg)						                            '매입확정여부          
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_IV_DT))	                                            '매입일                
		iRowStr = iRowStr & Chr(11) & ConvSPChars(Ucase(rs0(M_IV_HDR_PUR_GRP)))	                                        '매입그룹              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PUR_GRP_NM))	                                        '매입그룹명            
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_IV_CUR))	                                            '화폐                  
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_NET_DOC_AMT), ggAmtOfMoney.DecPoint,0)	    '매입순금액            	
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_TOT_VAT_DOC_AMT), ggAmtOfMoney.DecPoint,0)    '부가세금액            
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_GROSS_DOC_AMT), ggAmtOfMoney.DecPoint,0)    '부가세금액    
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_VAT_TYPE))										    '부가세유형            
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_VAT_TYPE_NM))									    '부가세유형명          
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_IV_HDR_VAT_RT),ggExchRate.DecPoint,6)			    '부가세율              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_METH))										    '결제방법              
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_METH_NM))									    '결제방법명            	         		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_MENT_TERM))										'가격조건              	              		
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_PAY_MENT_TERM_NM))									'가격조건명            	    
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_IV_HDR_GL_NO))	
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             

		If iLoopCount - 1 < C_SHEETMAXROWS_D Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount-1) = istrData	
		   istrData = ""
		Else
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If
		
		rs0.MoveNext
	Loop
	

	istrData = Join(PvArr, "")

	intARows = iLoopCount
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
    SetConditionData = false
         
    
	If Not(rs1.EOF Or rs1.BOF) Then
		SoCompanyNm = rs1("BP_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtSoCompanyCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "수주법인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    exit function
		End If
	End If   		
 

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(1,7)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 

	strVal = ""
    UNISqlId(0) = "MM111QA101"
    UNISqlId(1) = "MM111MA103"		'수주법인조회 
    
	UNIValue(1,0) = "'zzzzzzzzzz'"            
    
    '수주법인조회 
    If Trim(Request("txtSoCompanyCd")) <> "" Then
		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	    UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,0) = "|"
	End If
	
    '계산서발행일 
    If Trim(Request("txtBillFrDt")) <> "" Then
		UNIValue(0,1) =  " '" & Trim(UniConvDate(Request("txtBillFrDt"))) & "' "	
    Else
        UNIValue(0,1) = "|"
	End If
			
    If Trim(Request("txtBillToDt")) <> "" Then
		UNIValue(0,2) =  " '" & Trim(UniConvDate(Request("txtBillToDt"))) & "' "	
    Else
        UNIValue(0,2) = "|"
	End If
	
    '매입일 
    If Trim(Request("txtIvFrDt")) <> "" Then
		UNIValue(0,3) =  " '" & Trim(UniConvDate(Request("txtIvFrDt"))) & "' "	
    Else
        UNIValue(0,3) = "|"
	End If
			
    If Trim(Request("txtIvToDt")) <> "" Then
		UNIValue(0,4) =  " '" & Trim(UniConvDate(Request("txtIvToDt"))) & "' "	
    Else
        UNIValue(0,4) = "|"
	End If
	
    '매입확정처리여부 
    If Trim(Request("rdoCfmflg")) <> "" Then
		UNIValue(0,5) = " '"& FilterVar(Trim(UCase(Request("rdoCfmflg"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,5) = "|"
	End If	
	
    '발주번호 
    If Trim(Request("txtCustPoNo")) <> "" Then
		UNIValue(0,6) = " '"& FilterVar(Trim(UCase(Request("txtCustPoNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,6) = "|"
	End If
	
    '공급처세금계산서번호 
    If Trim(Request("txtBlNo")) <> "" Then
		UNIValue(0,7) = " '"& FilterVar(Trim(UCase(Request("txtBlNo"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,7) = "|"
	End If			

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
         
	If SetConditionData = False Then Exit Sub
	

    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
            Call  MakeSpreadSheetData()

    End If  
End Sub


'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function


%>
