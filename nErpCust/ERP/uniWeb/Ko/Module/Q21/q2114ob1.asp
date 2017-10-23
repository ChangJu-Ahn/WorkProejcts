<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<% Call loadInfTB19029B("I", "*", "NOCOOKIE","MB") %>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2114OB1
'*  4. Program Name         : 이력카드출력 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd
Dim strPlantCd 
Dim strInspClassCd 
Dim strItemCd 
Dim i
Dim var2
Dim intRecordCount1
Dim intRecordCount2

strPlantCd = FilterVar(Trim(Request("txtPlantCd")), "''", "SNM")
strInspClassCd = FilterVar(Trim(Request("txtInspClassCd")), "''", "SNM")
strItemCd = FilterVar(Trim(Request("txtItemCd")), "''", "SNM")

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	Dim lgStrSQL
	    
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL = "Select b.insp_item_cd from q_Inspection_standard_by_item a, q_inspection_item b " &_
		"where  a.insp_item_cd = b.insp_item_cd and a.plant_cd =  " & FilterVar(strPlantCd , "''", "S") & " and a.item_cd =  " & FilterVar(strItemCd , "''", "S") & " and a.insp_class_cd =  " & FilterVar(strInspClassCd , "''", "S") & ""	
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then
		intRecordCount1 = 0
		intRecordCount2 = 1
		Do While Not lgObjRs.EOF
			var2= var2 & "InspItemCd" & CStr(intRecordCount2) & "|" & ConvSPChars(lgObjRs(0)) & "|"		
			lgObjRs.MoveNext
			intRecordCount1 = intRecordCount1 + 1
			intRecordCount2 = intRecordCount2 + 1
		Loop

		For i = (intRecordCount1 + 1) to 16
			var2= var2 & "InspItemCd" & CStr(i) & "|" & "zzzz" & "|"
		Next		
		
		Call SubCloseRs(lgObjRs)			'☜ : Release RecordSSet
	Else
		Call DisplayMsgBox("229902", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
		Response.End
	End If
End Sub 
	
Call SubOpenDB(lgObjConn)		'☜: Make a DB Connection
Call SubBizQuery()				'☜: Query
Call SubCloseDB(lgObjConn)      '☜: Close DB Connection
%>
<Script language=VbScript>
Parent.strInspItemCd = "<%=var2%>"
Call Parent.DbQueryOK
</script>
