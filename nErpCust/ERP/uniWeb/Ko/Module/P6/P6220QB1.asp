<%@LANGUAGE = VBScript%>
<% Option Explicit%>
<%
'***********************************************************************************************************************
'*  1. Module Name          : M8
'*  2. Function Name        : 매입일보(HB)
'*  3. Program ID           : XM007OB_KO321
'*  4. Program Name         : 매입일보(HB)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/05/03
'*  8. Modified date(Last)  : 2005/06/28
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
	lgStrSQL = "select a.fac_cast_cd as ab,b.cast_nm as bc,a.work_dt as cd,dbo.ufn_GetCodeName('Z425', d.zinsp_part) as de,a.insp_text as ef "
	lgStrSQL = lgStrSQL & " ,g.bp_nm as fg ,dbo.ufn_H_GetEmpName(f.insp_emp_cd) as gh "
	lgStrSQL = lgStrSQL & " from y_fac_cast_plan as a,y_cast as b,y_fac_cast_check as d "
	lgStrSQL = lgStrSQL & " ,y_fac_cast_repair as f,b_biz_partner as g"
	lgStrSQL = lgStrSQL & " where a.fac_cast_cd = b.cast_cd"
	lgStrSQL = lgStrSQL & " and	a.gubun_cd='20' "
	lgStrSQL = lgStrSQL & " and	a.plan_gubun='2'"
	lgStrSQL = lgStrSQL & " and a.gubun_cd=d.gubun_cd"
	lgStrSQL = lgStrSQL & " and	a.work_dt = d.work_dt"
	lgStrSQL = lgStrSQL & " and a.plan_gubun = d.plan_gubun"
	lgStrSQL = lgStrSQL & " and	a.fac_cast_cd = f.fac_cast_cd"
	lgStrSQL = lgStrSQL & " and	a.work_dt = f.work_dt"
	lgStrSQL = lgStrSQL & " and	a.plan_gubun = f.plan_gubun "
	lgStrSQL = lgStrSQL & " and	f.cust_cd *= g.bp_cd "
		
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
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ef"))
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("fg"))
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
	Call SubCloseDB(lgObjConn)  
	Response.End																				'☜: Process End

End Select

%>
