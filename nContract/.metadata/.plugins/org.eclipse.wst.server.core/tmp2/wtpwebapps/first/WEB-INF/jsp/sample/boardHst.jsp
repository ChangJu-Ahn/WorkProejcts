<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%@ include file="/WEB-INF/include/include-header.jspf" %>
</head>
<body>
	<jsp:include page="../sample/top.jsp" flush="false"></jsp:include>
	<H2>체결 계약서 변경 이력</H2>
    <table class="board_view">
			<colgroup>
				<col width="15%"/>
				<col width="25%"/>
				<col width="15%"/>
				<col width="15%"/>
				<col width="15%"/>
				<col width="15%"/>
			</colgroup>
			<tbody>
				<tr>
					<th scope="row">계약번호</th>
					<td colspan="5">${map.CONTRACT_NO }</td>
				</tr>
				<tr>
					<th scope="row">사업부</th>
					<td colspan="2">${map.BU_CODE }</td>
					<th scope="row">해지조건</th>
					<td colspan="2">${map.EXPIRE_CONDITION }</td>
				</tr>
				<tr>
					<th scope="row">고객사(업체)명</th>
					<td colspan="2">${map.BP_CD1 }</td>
					<th scope="row">구분</th>
					<td colspan="2">${map.BP_TYPE }</td>
				</tr>
				<tr>
					<th scope="row">고객사(업체)명2</th>
					<td colspan="2">${map.BP_CD2 }</td>
					<th scope="row">계약상태</th>
					<td colspan="2">
						<c:choose>
                       		<c:when test="${map.EXPIRE_FLAG == 'Y' && map.MODIFY_FLAG == 'Y'}">
                       			해지 / 해지일 :${map.END_DT }
                       		</c:when>
                       		<c:when test="${map.EXPIRE_FLAG == 'N' && map.MODIFY_FLAG == 'Y'}">
                       			변경
                       		</c:when>
                       		<c:otherwise>
                       			등록
                       		</c:otherwise>
	                       </c:choose>
	                </td>
				</tr>
				<tr>
					<th scope="row">계약서 명</th>
					<td colspan="5">${map.CONTRACT_NM }
						<input type="hidden" id="CONTRACT_NO" name ="CONTRACT_NO" value="${map.CONTRACT_NO }">
					</td>
				</tr>
				<tr>
					<th scope="row">계약 구분</th>
					<td>${map.CONTRACT_TYPE }</td>
					<th scope="row">목적사업(제품)</th>
					<td colspan="3">${map.PURPOSE }</td>
				</tr>
				<tr>
					<th scope="row">효력발생일</th>
					<td>${map.VALID_DT }</td>
					<th scope="row">기간만료일</th>
					<td>${map.EXPIRE_DT }</td>
					<th scope="row">자동연장</th>
					<td>${map.EXTEND_FLAG } 
						<c:if test="${not empty map.EXTEND_TERM }">
						/ ${map.EXTEND_TERM }
						</c:if>
					</td>
				</tr>
				<tr>
					<th scope="row">부속계약서</th>
					<td colspan="3">${map.P_CONTRACT }</td>
					<th scope="row">해지통지기간</th>
					<td>${map.EXPIRE_TERM }</td>
				</tr>
				<tr>
					<th scope="row">비고</th>
					<td colspan="5">${map.REMARK }</td>
				</tr>
				<tr>
					<th scope="row"> 변경 내용</th>
					<td colspan="5">${map.MODIFY_CONTENT }</td>
				</tr>
				<tr>
					<th scope="row">파일</th>
					<td colspan="5">
						<c:forEach var="row" items="${list }">
	                       <div>
	                       		<input type="hidden" id="SEQ" value="${row.SEQ }">
	                       		<a href="#this" name="file">${row.ORIGINAL_FILE_NAME }</a> 
	                       		(${row.FILE_SIZE }kb)
	                       <br/>
	                       </div>
                    	</c:forEach>	
					</td>
				</tr>
				<tr>
					<th scope="row">기타</th>
					<td colspan="5">${map.SYS_REMARK }</td>
				</tr>
			</tbody>
	</table>
	
	    <!-- Role 불러 오기 -->
    <sec:authentication property="principal.username"  var ="currentUsername"/>
    <sec:authorize access="hasRole('ROLE_ADMIN')" var="ROLE_ADMIN"></sec:authorize>
    
    <div style="width: 60%;-style: inside; text-align: left;margin: auto; padding-bottom: 50px;">
    <br/>
		<c:if test="${map.INSERT_USER == currentUsername || ROLE_ADMIN == true}"> <!-- 자기가 작성했거나, 운영자만 볼 수 있음 -->
			<c:if test="${map.MAIN_EXPIRE_FLAG != 'Y'}">						  <!-- 최종 계약서의 만료상태를 확인하여 만료가 되지 않은 계약서만 수정가능 -->
				<a href="#this" class="btn" id="delete">삭제하기</a>
				<a href="#this" class="btn" id="update">수정하기</a>
   			</c:if>
  		</c:if>
   		<a href="#this" class="btn" id="list">목록으로</a>
    </div>
     

    <%@ include file="/WEB-INF/include/include-body.jspf" %>
    <script type="text/javascript">
    	var CONTRACT_NO = "${map.CONTRACT_NO }";
    	var HST_SEQ = "${map.HST_SEQ}";
    	
        $(document).ready(function(){
            //목록으로 버튼
        	$("#list").on("click", function(e){ 
                e.preventDefault();
                fn_openBoardList();
            });
            
        	//파일 이름
            $("a[name='file']").on("click", function(e){ 
                e.preventDefault();
                fn_downloadFile($(this));
            });
            
            //삭제하기 버튼
            $("#delete").on("click", function(e){
            	Lobibox.confirm({
            		msg: "한번 삭제 된 정보는 다시 복원할 수 없습니다. 삭제하시겠습니까?",
        		    callback: function ($this, type, ev) {
        		        if(type == "yes"){
        		        	e.preventDefault();
        	             	fn_openBoardDelete();
        		        }
        		    }
            	});
            });
            
            //수정하기 버튼
            $("#update").on("click", function(e){
//             	Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
//           	    		 {
//           	    		     msg: "기능구성 중 입니다."
//           	    		 });
//            	return;
            	
            	e.preventDefault();
            	fn_openBoardUpdate();
            });
        });
         
        function fn_openBoardList(){
            var comSubmit = new ComSubmit();
            
            comSubmit.setUrl("<c:url value='/sample/openBoardList.do' />");
            comSubmit.submit();
        }
        
        function fn_openBoardDelete(){
            var comSubmit = new ComSubmit();
            
            comSubmit.setUrl("<c:url value='/sample/BoardHstDelete.do'/>");
            
            comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
            comSubmit.addParam("HST_SEQ", HST_SEQ);
            comSubmit.addParam("userid", "${currentUsername}");
            
            comSubmit.submit();
        }
        
        function fn_openBoardUpdate(){
            var comSubmit = new ComSubmit();
            
            comSubmit.setUrl("<c:url value='/sample/openBoardHstUpdate.do' />");
            
            comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
            comSubmit.addParam("HST_SEQ", HST_SEQ);
            
            comSubmit.submit();
        }
           
        function fn_downloadFile(obj){
	        var CONTRACT_NO = $("#CONTRACT_NO").val();
	        var SEQ = obj.parent().find("#SEQ").val();
	        var comSubmit = new ComSubmit();
	        
	        comSubmit.setUrl("<c:url value='/common/downloadFile.do' />");
	        
	        comSubmit.addParam("CONTRACT_NO", CONTRACT_NO);
	        comSubmit.addParam("SEQ", SEQ);
	        
	        comSubmit.submit();
    	}
    </script>
</body>
</html>