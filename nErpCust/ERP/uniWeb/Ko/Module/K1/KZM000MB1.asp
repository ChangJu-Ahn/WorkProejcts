<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : ZM000MA1
'*  4. Program Name         : 멀티컴퍼니접속정보등록 
'*  5. Program Desc         : 멀티컴퍼니접속정보등록  
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/07/02
'*  8. Modified date(Last)  : 2005/04/07
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : Moon Jung Gil
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
Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows

intARows=0
intTRows=0
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status


Call HideStatusWnd                                                               '☜: Hide Processing message
lgOpModeCRUD  = Request("txtMode") 

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
		 Call  SubBizQueryMulti()
    Case CStr(UID_M0002)                                                         '☜: Save,Update
         Call SubBizSaveMulti()
    Case CStr(UID_M0003)
         Call SubBizSaveMulti()
End Select

Sub SubBizQueryMulti()


	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

'	Call DisplayMsgBox(lgStrSQL, vbInformation, "", "", I_MKSCRIPT)


	Call FixUNISQLData()		'☜ : DB-Agent로 보낼 parameter 데이타 set
	
	Call QueryData()			'☜ : DB-Agent를 통한 ADO query
	
	'-----------------------
	'Result data display area
	'----------------------- 

%>

	<Script Language=vbscript>
		With parent
			.frm1.txtCompanyCd.value = "<%=ConvSPChars(Request("txtCompanyCd"))%>"			
			.frm1.txtCompanyNm.Value	= "<%=SoCompanyNm%>"							
			.frm1.txtCompanyCd.focus
			
			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then
				
				'Show multi spreadsheet data from this line
				       
				.ggoSpread.Source    = .frm1.vspdData 
				.ggoSpread.SSShowData "<%=istrData%>"                  '☜: Display data 
				
				.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
				
				.DbQueryOk '<%=intARows%>,<%=intTRows%>
							
			End If
		End with
	</Script>	
<%	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
 
    Dim ZMG000
    Dim iErrorPosition
	Dim txtSpread
								
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                      '☜: Clear Error status


    
    Set ZMG000 = Server.CreateObject("PZMG000.cZMcUrlInfo") 

    txtSpread = Trim(Request("txtSpread"))
         
    Call  ZMG000.Z_MC_URL_INFO_SVR(gStrGlobalCollection, txtSpread, iErrorPosition)	  
 
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
		Set ZMG000 = Nothing												'☜: ComProxy Unload
		exit sub
 	end if
        
    Set ZMG000 = Nothing	
                 
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DBSaveOK "           & vbCr
    Response.Write "</Script>"                  & vbCr 
               
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

		Const		Z_MC_URL_INFO_COMPANY_CD	=	0
		Const		B_BIZ_PARTNER_BP_NM			=	1
		Const		Z_MC_URL_INFO_MC_FLG		=	2
		Const		Z_MC_URL_INFO_URL_TXT	    =	3
		Const		Z_MC_URL_INFO_USE_FLG	    =	4

				    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
	
	'----- 레코드셋 칼럼 순서 ----------
	'A.COMPANY_CD, B.BP_NM, A.MC_FLG, A.URL_TXT, A.USE_FLG
	'-----------------------------------

	iLoopCount = 0
    ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""

		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(Z_MC_URL_INFO_COMPANY_CD))	
		iRowStr = iRowStr & Chr(11) & ""				
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_BIZ_PARTNER_BP_NM)) 
        If ConvSPChars(rs0(Z_MC_URL_INFO_MC_FLG))  = "M" then
           iRowStr = iRowStr & Chr(11) & "발주"
        Else
        	iRowStr = iRowStr & Chr(11) & "수주"
        End If			    					
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(Z_MC_URL_INFO_URL_TXT)) 
        If ConvSPChars(rs0(Z_MC_URL_INFO_USE_FLG))  = "Y" then
           iRowStr = iRowStr & Chr(11) & "1"
        Else
        	iRowStr = iRowStr & Chr(11) & "0"
        End If										 		
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
		If Len(Request("txtCompanyCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "거래법인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
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
    Redim UNIValue(1,1)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 

	strVal = ""
    UNISqlId(0) = "ZM000MA1"
    UNISqlId(1) = "MM111MA103"		'법인조회 
    
	UNIValue(1,0) = "'zzzzzzzzzz'"            

    '법인조회 
    If Trim(Request("txtCompanyCd")) <> "" Then
	    UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtCompanyCd"))), " " , "SNM") & "' "
	    UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtCompanyCd"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,0) = "|"
	End If
	
    '사용여부 
    If Trim(Request("rdoUsageFlg")) <> "" Then
		UNIValue(0,1) = " '"& FilterVar(Trim(UCase(Request("rdoUsageFlg"))), " " , "SNM") & "' "
	Else 
	    UNIValue(0,1) = "|"
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
