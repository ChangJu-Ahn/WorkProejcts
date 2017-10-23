<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(User Defined Major Code)
'*  3. Program ID           : b1a03mb1.asp
'*  4. Program Name         : b1a03mb1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/07/06
'*  7. Modified date(Last)  : 2002/12/12
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next								'☜: 

Dim pB1A031											'입력/수정용 ComProxy Dll 사용 변수 
Dim pB1A038											'조회용 ComProxy Dll 사용 변수 

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread

Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow
Dim GroupCount          

'@Var_Declare

Call LoadBasisGlobalInf()

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strSpread = Trim(Request("txtSpread"))

On Error Resume Next

Select Case strMode

Case CStr(UID_M0001)								'☜: 현재 조회/Prev/Next 요청을 받음 

    Dim I1_b_user_defined_major
    Const B392_I1_ud_major_cd = 0 
    Const B392_I1_ud_major_nm = 1
 
    Const B392_EG1_E1_ud_major_cd = 0 
    Const B392_EG1_E1_ud_major_nm = 1
    Const B392_EG1_E1_ud_minor_len = 2

	Dim ObjPB2G191
	Dim Export_Array

    ReDim I1_b_user_defined_major(B392_I1_ud_major_nm)
%>
<Script Language=vbscript>
	parent.frm1.txtMajorNm.value = "<%=ConvSPChars(LookUpMajor(Request("txtMajor")))%>"
</Script>
<%  
	I1_b_user_defined_major(B392_I1_ud_major_cd) = Request("txtMajor")
	
    Set ObjPB2G191 = server.CreateObject ("PB2G191.cBListUserMajorCode")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G191.B_LIST_USER_MAJOR_CODE (gStrGlobalCollection,I1_b_user_defined_major)
    Set ObjPB2G191 = nothing

    If CheckSYSTEMError(Err,True) = True Then                               
		Response.End														'☜: 비지니스 로직 처리를 종료함 
    End If
    on error goto 0
%>

<Script Language=vbscript>
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp
    Dim strData
	
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
		
		 LngMaxRow = 0
<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>        
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B392_EG1_E1_ud_major_cd)))%>"'
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B392_EG1_E1_ud_major_nm)))%>"'
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B392_EG1_E1_ud_minor_len)))%>"'
            strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
            strData = strData & Chr(11) & Chr(12)
<%      		
        Next
%>    		
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData
		
		.frm1.hMajorCd.value = "<%=ConvSPChars(Request("txtMajor"))%>"
		.DbQueryOk
		
	End With
</Script>	
<%    
    
Case CStr(UID_M0002)																'☜: 저장 요청을 받음 
									
    Dim Obj2PB2G191

    Set Obj2PB2G191 = server.CreateObject ("PB2G191.cBCtlUserMajorCode")    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G191.B_CONTROL_USER_MAJOR_CODE(gStrGlobalCollection,strSpread)
    Set Obj22PB2G191 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
		Response.End 
    End If
    on error goto 0                                                                  '☜: Unload Comproxy

%>
<Script Language=vbscript>
	With parent																	    '☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
<%					

End Select

%>

<%
'==============================================================================
' Function : LookUp...
' Description : 저장시 Lookup
'==============================================================================
Function LookUpMajor(Byval strCode)
    Const B393_I1_ud_major_cd = 0
    Const B393_I1_ud_major_nm = 1

    Const B393_E1_ud_major_cd = 0
    Const B393_E1_ud_major_nm = 1
    Const B393_E1_ud_minor_len = 2
    
	Dim ObjPB2G192	
	Dim I1_b_user_defined_major
	Dim E1_b_major
	
    ReDim I1_b_user_defined_major(B393_I1_ud_major_nm)
    ReDim E1_b_major(B393_E1_ud_minor_len)
    

    I1_b_user_defined_major(B393_I1_ud_major_cd) = Request("txtMajor")
	
    Set ObjPB2G192 = server.CreateObject ("PB2G192.cBLookUserMajorCode")    
    On Error Resume Next                                                                 
    Err.Clear                                                                            
    E1_b_major = ObjPB2G192.B_LOOKUP_USER_MAJOR_CODE(gStrGlobalCollection,I1_b_user_defined_major)
    Set ObjPB2G192 = nothing    
    
    If Err.number <> 0 and inStr(Err.Description ,"122400") > 0 then
  	LookUpMajor = ""
    Else
        If CheckSYSTEMError(Err,True) = True Then                                              
        	Exit Function
	    End If
        on error goto 0

	    LookUpMajor = E1_b_major(B393_E1_ud_major_nm)
    End If
					  
End Function
%>

