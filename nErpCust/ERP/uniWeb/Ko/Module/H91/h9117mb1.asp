<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Const C_SHEETMAXROWS_D = 100
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("Q", "H", "NOCOOKIE", "MB")

	Dim strYear, strFrDt, strToDt
	    
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
'    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    
	strYear		= FilterVar(left(lgKeyStream(5),4), "''", "S")
	strFrDt		= FilterVar(left(lgKeyStream(5),4) & "0101", "''", "S")
	strToDt		= FilterVar(left(lgKeyStream(5),4) & "1231", "''", "S")
		
 ON ERROR RESUME NEXT
 
     Call SubCreateCommandObject(lgObjComm)
     Call SubBizQuery()
     Call SubCloseCommandObject(lgObjComm)

 '============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt
	Dim sDstbOrder, sSenderCostCd, sFromAcctCd, sToAcctCd, sType, sSendAmt, sRecvAmt, arrTmp
    Dim DFnm 

    Call SubCreateCommandObject(lgObjComm)	

    With lgObjComm
		.CommandTimeout = 0

		.CommandText = "dbo.usp_h_yearend_hometax"  
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@BIZ_AREA_CD",	adVarXChar,	adParamInput, 10, lgKeyStream(1) )	'신고사업장 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SEND_DT",		adVarXChar,	adParamInput, 8,  lgKeyStream(2) )	'제출연월일 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@BASE_DT",		adVarXChar,	adParamInput, 8,  lgKeyStream(3) )	'기준연월일 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ALL_YN",		adVarXChar,	adParamInput, 1,  lgKeyStream(4) )	'통합신고여부 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@RETIRE_YN",		adVarXChar,	adParamInput, 1,  lgKeyStream(5) )	'퇴직포함여부 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@GUBUN",			adVarXChar,	adParamInput, 1,  lgKeyStream(6) )	'제출자구분 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@GIGAN",			adVarXChar,	adParamInput, 1,  lgKeyStream(7) )	'대상기간 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SEMU",			adVarXChar,	adParamInput, 6,  lgKeyStream(8) )	'세무대리인관리번호 

        Set oRs = lgObjComm.Execute

    End With

'-----------------------------------------------------------------
' SP에서 A,B,C,D,E 레코드 동시에 RETURN
 
    If Not oRs.EOF Then

' ------------- A 레코드 
		Dim arrColRowA, i, j
	 
		arrColRowA = oRs.GetRows()
		iLngRowCnt	= UBound(arrColRowA, 2) 
		iLngColCnt	= UBound(arrColRowA, 1) 

		For i = 0 To iLngRowCnt
					
			For j = 0 To  iLngColCnt 
				lgstrData = lgstrData & Chr(11) & arrColRowA(j, i) 
						
			Next
					
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + j
			lgstrData = lgstrData & Chr(11) & Chr(12)
		Next

		Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
' Response.Write "sTxt:" & sTxt

' -------------B 레코드 
		If oRs.EOF = TRUE Then
 			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정		
		Else
 			Dim arrColRowB
			arrColRowB		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			iLngRowCnt	= UBound(arrColRowB, 2) 
			iLngColCnt	= UBound(arrColRowB, 1)
		
 			For i = 0 To iLngRowCnt

				For j = 0 To  iLngColCnt 
					lgstrData1 = lgstrData1 & Chr(11) & arrColRowB(j, i) 
						
				Next
					
				lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + j
				lgstrData1 = lgstrData1 & Chr(11) & Chr(12)
			Next
  		End If
' ------------- C 레코드 
		If oRs.EOF = TRUE Then
 			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정		
		Else

 			Dim arrColRowC
			arrColRowC		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			iLngRowCnt	= UBound(arrColRowC, 2) 
			iLngColCnt	= UBound(arrColRowC, 1)
		
 			For i = 0 To iLngRowCnt

				For j = 0 To  iLngColCnt 
					lgstrData2 = lgstrData2 & Chr(11) & arrColRowC(j, i) 
						
				Next
					
				lgstrData2 = lgstrData2 & Chr(11) & lgLngMaxRow + j
				lgstrData2 = lgstrData2 & Chr(11) & Chr(12)
			Next
 		End If
 ' ------------- D 레코드 

		If oRs.EOF = TRUE Then
 			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정		
		Else

 			Dim arrColRowD
			arrColRowD		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			iLngRowCnt	= UBound(arrColRowD, 2) 
			iLngColCnt	= UBound(arrColRowD, 1)
		
 			For i = 0 To iLngRowCnt

				For j = 0 To  iLngColCnt 
					lgstrData3 = lgstrData3 & Chr(11) & arrColRowD(j, i) 
						
				Next
					
				lgstrData3 = lgstrData3 & Chr(11) & lgLngMaxRow + j
				lgstrData3 = lgstrData3 & Chr(11) & Chr(12)
			Next
		End If
dim tRow,tmp,lgstrData4a
 ' ------------- E 레코드 
		If oRs.EOF = TRUE Then
 			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정		
		Else 
 			Dim arrColRowE
			arrColRowE		= oRs.GetRows()
		
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			iLngRowCnt	= UBound(arrColRowE, 2) 
			iLngColCnt	= UBound(arrColRowE, 1)

 			For i = 0 To iLngRowCnt


				
				'tRow=	ubound(split( arrColRowE(0,i),chr(11)))
			
				'if cint(tRow)<>93 then
				'	for j=0 to 93-tRow-1
				'		tmp = tmp & chr(11)
				'	next
				'end if	
			
			 tmp = split(arrColRowE(0,i),chr(11))
				 
				 for j=0 to uBound(tmp)-1
					if IsNumeric(tmp(j)) then
						lgstrData4a = lgstrData4a & cDbl(tmp(j)) & chr(11)
					else
						lgstrData4a = lgstrData4a &  tmp(j) & chr(11)
					end if
				 next
				 		
			lgstrData4 = lgstrData4 & "" &CHR(11) & lgstrData4a  & chr(12)
				
			lgstrData4a=""

			Next
		End If
				
		Dim li_biz_own_rgst_no  

        li_biz_own_rgst_no = Trim(lgKeyStream(0))  

        If Trim(li_biz_own_rgst_no) = "" Or Trim(li_biz_own_rgst_no) <> Trim(replace(arrColRowA(8,0),"-","")) Then 
            li_biz_own_rgst_no = Trim(replace(arrColRowA(8,0),"-",""))
            li_biz_own_rgst_no = Left(li_biz_own_rgst_no,7) & "." & Right(li_biz_own_rgst_no,3)
        End If

'		If Trim(lgCurrentSpd) = "A" then
			DFnm = "C:\c" & li_biz_own_rgst_no       
            
%>
<SCRIPT LANGUAGE=VBSCRIPT>
		parent.frm1.txtFile.value = "<%=DFnm%>"
</SCRIPT>
<%      
'		End If
	
    Else 
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If       
End Sub	
%>


<Script Language="VBScript">

    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
              
                .ggoSpread.Source     = .frm1.vspdData
				.ggoSpread.SSShowData "<%=lgstrData%>"
 
                .ggoSpread.Source     = .frm1.vspdData1
				.ggoSpread.SSShowData "<%=lgstrData1%>"

                .ggoSpread.Source     = .frm1.vspdData2
				.ggoSpread.SSShowData "<%=lgstrData2%>"

                .ggoSpread.Source     = .frm1.vspdData3
				.ggoSpread.SSShowData "<%=lgstrData3%>"

                .ggoSpread.Source     = .frm1.vspdData4
				.ggoSpread.SSShowData "<%=lgstrData4%>"									
                End with
          End If   
    End Select    
       
</Script>	
