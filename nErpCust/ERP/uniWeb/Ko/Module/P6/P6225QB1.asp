<%@LANGUAGE = VBScript%>
<% Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : p6
'*  2. Function Name        : 금형점검내역조회(HB)
'*  3. Program ID           : XP008OA_KO321
'*  4. Program Name         : 금형점검내역조회(HB)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/05/03
'*  8. Modified date(Last)  : 2005/07/20
'*  9. Modifier (First)     : Yoo Myung Sik
'* 10. Modifier (Last)      : Joo Young Hoon
'* 11. Comment              :
'************************************************************************************************************************

%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->	
<!-- #Include file="../../inc/adovbs.inc" -->	
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 


    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)


Call HideStatusWnd

Dim strMode		
Dim LngMaxRow							' 현재 그리드의 최대Row
Dim LngRow
Dim StrSeq
Dim lsSEQ
Dim lsCnt1
 
Dim txtCastCd, lsReqdlvyFromDt, lsReqdlvyToDt

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

Select Case strMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
    
    Err.Clear                                                                '☜: Protect system from crashing

	txtCastCd = trim(Request("txtCastCd"))
	lsReqdlvyFromDt = Request("txtReqdlvyFromDt")
	lsReqdlvyToDt = Request("txtReqdlvyToDt")
	
	
	'설비명 조회 
	If Trim(txtCastCd) <> "" Then
		lgStrSQL = "SELECT FACILITY_NM FROM Y_FACILITY (NOLOCK) WHERE FACILITY_CD = " & Filtervar(txtCastCd, "''", "S")      
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write " <Script Language=""VBScript""> " & vbCrLf
			Response.Write " Parent.frm1.txtCastNm.value =  """"" & vbCrLf
			Response.Write " </Script> " & vbCrLf	
		Else
			Response.Write " <Script Language=""VBScript""> " & vbCrLf
			Response.Write " Parent.frm1.txtCastNm.value =  """ & ConvSPChars(lgObjRs("FACILITY_NM")) & """" & vbCrLf
			Response.Write " </Script> " & vbCrLf
		End if   
	Else
		Response.Write " <Script Language=""VBScript""> " & vbCrLf
		Response.Write " Parent.frm1.txtCastNm.value =  """"" & vbCrLf
		Response.Write " </Script> " & vbCrLf	   
	End If
	    
	'메인 조회 
	lgStrSQL = "select a.fac_cast_cd as ab,b.cast_nm as bc,a.work_dt as cd "
	lgStrSQL = lgStrSQL & " ,a.insp_text as de,isnull(insp_hour,0) as ef,isnull(insp_min,0) as fg,insp_dept as gh "
	lgStrSQL = lgStrSQL & " from y_fac_cast_plan as a,y_cast as b"
	lgStrSQL = lgStrSQL & " where a.fac_cast_cd = b.cast_cd"
	lgStrSQL = lgStrSQL & " and a.gubun_cd='20'"
	lgStrSQL = lgStrSQL & " and plan_gubun='1' "
				
	if Trim(lsReqdlvyFromDt) <> "" Then 
		lgStrSQL = lgStrSQL & " and a.work_dt >='" & lsReqdlvyFromDt & "' " 
	end if
		
	if Trim(lsReqdlvyToDt) <> "" Then 
		lgStrSQL = lgStrSQL & " and a.work_dt <='" & lsReqdlvyToDt & "' " 
	end if
		
	if Trim(txtCastCd) <> "" Then 
		lgStrSQL = lgStrSQL & " and a.fac_cast_cd = '" & txtCastCd & "' " 
	end if
	
	
    lgStrSQL = lgStrSQL & " order by ab"       
      
    


    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    else 
		lsCnt1 = 0

        Do While Not lgObjRs.EOF

						
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ab"))
		    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bc"))
	        lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("cd"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("de"))
	        lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("ef"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
	        lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("fg"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gh"))
			lgstrData = lgstrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext
             
        Loop 


    End if 
 	
 	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                              '☜: Release RecordSSet



%>
    
<Script Language="VBScript">

             With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"

                .DBQueryOk                  
	         End with
       
</Script>	

<%
																				'☜: Process End
																				 '☜: Process End
	Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
	
	Response.End
	
End Select	
%>
