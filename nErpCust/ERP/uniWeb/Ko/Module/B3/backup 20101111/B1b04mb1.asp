<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : B1b04mb1
'*  4. Program Name         : HS부호등록 
'*  5. Program Desc         : HS부호등록 
'*  6. Component List       : PB1GB48.cBListHsCodeS
'*  7. Modified date(First) : 2000/04/18
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : An Chang Hwan
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<%
	On Error Resume Next
	Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
	Call HideStatusWnd

	Dim strMode																		'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

	strMode = Request("txtMode")		

	Select Case strMode
		Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
			Call SubBizQueryMulti()
			
	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	Dim iB1b048																' OpenBpCd L/C Detail 조회용 Object
	Dim iLngRow
	Dim LngMaxRow
	Dim intGroupCount
	Dim lgStrPrevKey
	Dim StrNextKey
	Dim istrData
	Dim C_Max_Count
	Dim I1_hs_cd
	Dim I2_hs_cd_next
	Dim iMax
	Dim PvArr

	Dim E1_b_hs_code
	Const M019_E1_hs_cd = 0    
	Const M019_E1_hs_nm = 1

	Dim E2_next_hs_cd
			
	Dim EG1_export_group
	Const M019_EG1_E1_hs_cd = 0
	Const M019_EG1_E1_hs_nm = 1
	Const M019_EG1_E1_hs_spec = 2
	Const M019_EG1_E1_hs_unit = 3
	Const M019_EG1_E1_fix_rate = 4
	Const M019_EG1_E1_moun_cd = 5
	Const M019_EG1_E2_unit_nm = 6    
	Const M019_EG1_E2_dimension = 7

	Const C_SHEETMAXROWS_D = 100

	Err.Clear																'☜: Protect system from crashing
	On Error Resume Next

	LngMaxRow = Request("txtMaxRows")
	I1_hs_cd = Request("txtHsCd")
	I2_hs_cd_next = Request("lgStrPrevKey")	
	'---------------------------------- HS Code Data Query ----------------------------------
	Set iB1b048 = Server.CreateObject("PB1GB48.cBListHsCodeS")
	If CheckSYSTEMError(Err,True) = true then 		
		Exit Sub
	End if

	Call iB1b048.B_LIST_HS_CODE_SVR(gStrGlobalCollection, CLng(C_SHEETMAXROWS_D), I1_hs_cd, I2_hs_cd_next, E1_b_hs_code, _
										E2_next_hs_cd, EG1_export_group)  

	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iB1b048 = Nothing												'☜: Complus Unload
		Exit Sub															'☜: Terminate Biz. Logic
	End if

	Set iB1b048 = Nothing												'☜: Complus Unload

	'-------------------------------
	'Result data display area
	'------------------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	If I2_hs_cd_next = "" Then
		Response.Write "Parent.frm1.txtHsNm.value = """ & ConvSPChars(E1_b_hs_code(M019_E1_hs_nm))	& """" & vbCr
	End If

	If IsEmpty(EG1_export_group) and LngMaxRow < 1  Then
		Response.Write "</Script>" & vbCr
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	Else     
		iMax = UBound(EG1_export_group,1)
		ReDim PvArr(iMax)
		For iLngRow = 0 To ubound(EG1_export_group,1)
			istrData = Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M019_EG1_E1_hs_cd)) 
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M019_EG1_E1_hs_nm))
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M019_EG1_E1_hs_spec))
			istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M019_EG1_E1_hs_unit))
			istrData = istrData & Chr(11) & ""
			istrData = istrData & Chr(11) & LngMaxRow + iLngRow
			istrData = istrData & Chr(11) & Chr(12)
	
			PvArr(iLngRow) = istrData
		Next
	    istrData = Join(PvArr, "")

		StrNextKey = ConvSPChars(E2_next_hs_cd)
		Response.Write "With Parent "               & vbCr
		
		Response.Write " .ggoSpread.Source = .frm1.vspdData"              & vbCr
		Response.Write " .lgQuery = True"					              & vbCr
		Response.Write " .ggoSpread.SSShowData     """ & istrData         & """" & vbCr
		
		Response.Write "   .frm1.vspdData.ReDraw = False	                   " & vbCr
		Response.Write "   .SetSpreadColor1 -1, -1	                               " & vbCr
		Response.Write "   .frm1.vspdData.ReDraw = True	                       " & vbCr
    
		Response.Write " .lgStrPrevKey           = """ & StrNextKey       & """" & vbCr
		Response.Write " If .frm1.vspdData.MaxRows < " & C_SHEETMAXROWS_D & " And .lgStrPrevKey <> """ & """" & " Then " & vbcr
		Response.Write " .DbQuery "		    	& vbCr 
		Response.Write " Else "		    		& vbCr 
		Response.Write " .frm1.txtHHsCd.value = """ & ConvSPChars(Request("txtHsCd"))	& """" & vbCr
		Response.Write " .DbQueryOk "		    & vbCr 
		Response.Write " End If "		    	& vbCr 
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
	End If
	
End Sub    

%>
