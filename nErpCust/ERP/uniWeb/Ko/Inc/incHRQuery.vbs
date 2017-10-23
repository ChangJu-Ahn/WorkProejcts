	  
Function FuncCodeName(intSW, MajorCd, MinorCd)
    Dim iSelectList
    Dim iFromList
    Dim iWhereList, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
    
    
    Select Case intSW
        Case 1                                                  ' B_MAJOR
              iSelectList = " MINOR_NM "
              iFromList   = " B_MINOR  "
              iWhereList  = " MAJOR_CD = " & FilterVar(MajorCd, "''", "S") & " AND MINOR_CD = " & FilterVar(MinorCd, "''", "S") 
              
        Case 2                                                  ' B_ACCT_DEPT  : 부서코드명 
              iSelectList = " DEPT_NM "
              iFromList   = " B_ACCT_DEPT  "
              If Trim(MinorCd) > "" Then
                 iWhereList  = " DEPT_CD    = " & FilterVar(MajorCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= " & FilterVar(MinorCd, "''", "S") & ")"
              Else
                 iWhereList  = " DEPT_CD    = " & FilterVar(MajorCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
              End If   
        Case 3                                                  ' B_COUNTRY : 국적 
              iSelectList = " COUNTRY_NM "
              iFromList   = " B_COUNTRY  "
              iWhereList  = " COUNTRY_CD = " & FilterVar(MinorCd, "''", "S")      

        Case 4                                                  ' B_COMPANY : 회사코드 
              iSelectList = " CO_NM "
              iFromList   = " B_COMPANY  "
              iWhereList  = " CO_CD = " & FilterVar(MinorCd, "''", "S")  
        Case 5                                                  ' 내부부서코드 
              iSelectList = " INTERNAL_CD "
              iFromList   = " B_ACCT_DEPT  "
              If Trim(MinorCd) > "" Then
                 iWhereList  = " DEPT_CD    = " & FilterVar(MajorCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= " & FilterVar(MinorCd, "''", "S") & ")"   
              Else
                 iWhereList  = " DEPT_CD    = " & FilterVar(MajorCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())" 
              End If 
	End Select

    If 	CommonQueryRs(iSelectList,iFromList,iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
        FuncCodeName = MinorCd
    Else
        lgF0 = Split(lgF0,Chr(11))
        FuncCodeName = lgF0(0)
    End If

End Function

Function FuncDeptName(DeptCd, OrgChangeDt, lgIntCd, DeptNm, IntCd)
    Dim iWhereList
    Dim strIntCd, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    DeptNm = ""
    IntCd = ""

    If  OrgChangeDt > "" Then
		iWhereList = " DEPT_CD = " &  FilterVar( DeptCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= " & FilterVar( OrgChangeDt, "''", "S")  & ")"
    Else
        iWhereList = " DEPT_CD = " &  FilterVar( DeptCd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
    End If   

    If 	CommonQueryRs(" DEPT_NM,INTERNAL_CD "," B_ACCT_DEPT ",iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
        FuncDeptName = -1	' 부서테이블에 없습니다.
	exit function
    end if

    lgF0 = Split(lgF0,Chr(11))
    lgF1 = Split(lgF1,Chr(11))

    strIntCd = Trim(Replace(lgIntCd, "%", ""))

    if (strIntCd = "") OR (LEFT(lgF1(0), Len(strIntCd)) <> strIntCd) then
        FuncDeptName= -2	' 권한이 없습니다.
    else
        DeptNm = Trim(lgF0(0))
        IntCd = Trim(lgF1(0))
	FuncDeptName= 0
    End If

End Function

Function FuncGetAuth(PgmId, UsrId, plgIntCd)
	dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    If 	CommonQueryRs(" INTERNAL_CD,AUTH_YN "," HZA010T ", " MNU_ID = " &  FilterVar( UCase(PgmId), "''", "S") & " AND USR_ID = " & FilterVar( UCase(UsrId), "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	    plgIntCd = "1"		' 권한테이블에 없을 경우 모든 권한을 갖는다.
        FuncAuth = 0		' 권한테이블에 없습니다.
	exit function
    End If

    lgF0 = Split(lgF0,Chr(11))
    lgF1 = Split(lgF1,Chr(11))	' AUTH_YN

    if	Trim(lgF1(0)) = "N" then	' 권한을 check 하지 않는다.
    	plgIntCd = "1"			' 권한을 check 하지 않을 경우 모든 권한을 갖는다.
    else
	plgIntCd = Trim(lgF0(0))
    end if

    FuncAuth= 0

End Function


Function FuncGetEmpInf2(pEmpNo,plgIntCd,pEmpName,pDeptNm,pRollPstn,pPayGrd1,pPayGrd2,pEntrDt,pIntCd)
    Dim ADF                                                                    '☜ : declaration Variable indicating ActiveX Data Factory
    Dim lgstrRetMsg                                                            '☜ : declaration Variable indicating Record Set Return Message
    Dim rs, rs0                              '☜ : declaration DBAgent Parameter 
    Dim strlgIntCd
    dim iErrCode, iErrDesc
	dim iSelectList,	iWhereList
    Err.Clear
    FuncGetEmpInf2 =  0

    pEmpName   = ""
    pDeptNm    = ""
    pRollPstn  = ""
    pPayGrd1   = ""
    pPayGrd2   = ""
    pEntrDt    = ""	
    pIntCd     = ""	

    iSelectList = " NAME,DEPT_NM,ROLL_PSTN,PAY_GRD1,PAY_GRD2,ENTR_DT,INTERNAL_CD "
    iWhereList = " EMP_NO = " & FilterVar( pEmpNo , "''", "S")  

    If 	CommonQueryRs(iSelectList," HAA010T ",iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
		lgF0 = Split(lgF0,Chr(11))			    
		pEmpName   = Trim(lgF0(0))

		lgF6 = Split(lgF6,Chr(11))
		pIntCd     = Trim(lgF6(0))		
    End If	

	If IsNull(pEmpName) or pEmpName="" Then
		FuncGetEmpInf2 = -1 
		Exit Function
	End If 	

    strlgIntCd = Trim(Replace(plgIntCd, "%", ""))
    If strlgIntCd="" Then
        strlgIntCd="1"      '권한이 필요 없는 프로그램에서 호출시 Default값 설정 
    End If

    if (strlgIntCd = "") OR (LEFT(pIntCd, Len(strlgIntCd)) <> strlgIntCd) then
        FuncGetEmpInf2 = -2	' 권한이 없습니다.
		exit function
    else	

		lgF1 = Split(lgF1,Chr(11))	
		pDeptNm    = Trim(lgF1(0))
	
		lgF2 = Split(lgF2,Chr(11))
		pRollPstn  = Trim(FuncCodeName(1,"H0002",Trim(lgF2(0))))
	
		lgF3 = Split(lgF3,Chr(11))
		pPayGrd1   = Trim(FuncCodeName(1,"H0001",Trim(lgF3(0))))
	
		lgF4 = Split(lgF4,Chr(11))
		pPayGrd2   = Trim(lgF4(0))
	
		lgF5 = Split(lgF5,Chr(11))
		pEntrDt    = Trim(lgF5(0))
	
        FuncGetEmpInf2 = 0	' 정상 
    end if

End Function

Function FuncGetTermDept(plgIntCd, pChngDt, rFrDept, rToDept)

    Dim iWhereList, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    rFrDept = ""
    rToDept = ""

    If  pChngDt > "" Then
		iWhereList = " INTERNAL_CD LIKE " & FilterVar( plgIntCd & "%", "''", "S") & " AND ORG_CHANGE_DT=(SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT <= " & FilterVar( pChngDt, "''", "S") & ")"
    Else
        iWhereList = " INTERNAL_CD LIKE " & FilterVar( plgIntCd & "%", "''", "S") & " AND ORG_CHANGE_DT=(SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"
    End If   

    If 	CommonQueryRs(" MIN(internal_cd), MAX(internal_cd) "," B_ACCT_DEPT ",iWhereList ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
        FuncGetTermDept = -1	' 부서테이블에 없습니다.
	exit function
    end if

    lgF0 = Split(lgF0,Chr(11))
    lgF1 = Split(lgF1,Chr(11))

    rFrDept = Trim(lgF0(0))
    rToDept = Trim(lgF1(0))

    FuncGetTermDept = 0

End Function

Function FuncLastMonthDay(pDate, rDate)

   Dim strDate1
   Dim strDate2

   strDate1 = Trim(Replace(pDate, gComDateType, ""))
   if strDate1 = "" then
      strDate2 = Year(Date) & gComDateType & Right("0" & Month(Date),2)
   else
      strDate2 = Mid(strDate1, 1, 4) & gComDateType
      strDate2 = strDate2 & Mid(strDate1, 5, 2)
      strDate2 = strDate2 & gComDateType & "01"
   end if

   rDate = DateAdd("D",-1, DateAdd("M",1,strDate2))

   FuncLastMonthDay = Day(rDate)

End Function

function HRAskPRAspName(Byval pPgmId)
	Dim iCalledAspName
	iCalledAspName = ""
	iCalledAspName = AskPRAspName(pPgmId)
	HRAskPRAspName = iCalledAspName
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, pPgmId, "x")
		HRAskPRAspName = "../../ComASP/PRNotFound.asp"
	end if
end function
