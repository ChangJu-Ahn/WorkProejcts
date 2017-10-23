package first.sample.controller;

import java.net.URLDecoder;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;

import org.apache.log4j.Logger;
/*import org.apache.xmlbeans.impl.xb.xmlconfig.Extensionconfig.*;*/
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.springframework.security.core.context.*;
import org.springframework.stereotype.Controller;
import org.springframework.security.core.userdetails.User;					//add, 2016.11.21			
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.context.request.RequestContextHolder;		//add, 2016.11.21
import org.springframework.web.context.request.ServletRequestAttributes;	//add, 2016.11.21

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import first.common.common.CommandMap;
import first.sample.service.SampleService;

@Controller
public class SampleController {
	Logger log = Logger.getLogger(this.getClass());
	String clientUserId;  
	
	@Resource(name = "sampleService") 
	private SampleService sampleService;
	
	@RequestMapping(value = "/downExcel.do")
	public ModelAndView excelTransform(CommandMap commandMap) throws Exception {
		String tempTarget = (String)commandMap.get("TARGET");
		ModelAndView mv = new ModelAndView("excelView");
		
		//get the date
		List<Map<String, Object>> excelList = sampleService.getAllData(commandMap.getMap());
	
		//select mv.addObject after to confirm variable 
		if(tempTarget.equals("CONTRACT_LIST"))
			mv.addObject("list", excelList);
		else
			mv.addObject("adminList", excelList);
		
		return mv;
	}
	
	/*
	 * @RequestMapping(value="/sample/openMain.do") public ModelAndView openMain() throws Exception{ ModelAndView mv = new ModelAndView("/sample/top"); return mv; }
	 */

	@RequestMapping(value = "/login/login_form.do")
	public ModelAndView openLogin(@RequestParam(value = "error", required = false) String error
									,@RequestParam(value = "change", required = false) String change
								) throws Exception {
		ModelAndView mv = new ModelAndView("/login/login_form");
		
		mv.addObject("error", error);
		mv.addObject("change", change);

		return mv;
	}

	@RequestMapping(value = "/report/reportView.do")
	public ModelAndView openCheckLogin() throws Exception {
		ModelAndView mv = new ModelAndView("/report/report_form");
		return mv;
	}

	@RequestMapping(value = "/admin/adminView.do")
	public ModelAndView openadminView() throws Exception {
		ModelAndView mv = new ModelAndView("/admin/admin_page");
		return mv;
	}

	@RequestMapping(value = "/admin/openAdmin_account.do")
	public ModelAndView openadminUserManagement() throws Exception {
		ModelAndView mv = new ModelAndView("/admin/admin_UserManager");
		return mv;
	}
	
	@RequestMapping(value = "/admin/admin_SysInfo.do")
	public ModelAndView openadminSystemManagement() throws Exception {
		ModelAndView mv = new ModelAndView("/admin/admin_SysManager");
		return mv;
	}

	@RequestMapping(value = "/admin/userInfoView.do")
	public ModelAndView openUserInfoView(@RequestParam(value = "page", required = false) String s_page,
			@RequestParam(value = "rows", required = false) String s_rows, CommandMap commandMap) throws Exception {

		ModelAndView mv = new ModelAndView("jsonView");

		int page = Integer.parseInt(s_page);
		int perPageRow = Integer.parseInt(s_rows);

		int total = 0;
		int records = 0;
		int start = 0;
		int end = 0;

		end = perPageRow * page;
		start = end - (perPageRow - 1);

		commandMap.put("START", start);
		commandMap.put("END", end);

		List<Map<String, Object>> list = sampleService.selectUserList(commandMap.getMap());

		records = (int) list.get(0).get("TOTAL_COUNT");

		if (records > 0)
			total = (int) Math.ceil((double) records / (double) perPageRow);
		else
			total = 0;

		mv.addObject("list", list);
		mv.addObject("page", page);
		mv.addObject("total", total);
		mv.addObject("records", records);

		return mv;
	}

	@RequestMapping(value = "/admin/gridView.do")
	public ModelAndView openJqgrid(@RequestParam(value = "CODE", required = false) String code,
			@RequestParam(value = "page", required = false) String s_page, @RequestParam(value = "rows", required = false) String s_rows,
			CommandMap commandMap) throws Exception {

		ModelAndView mv = new ModelAndView("jsonView");

		int page = Integer.parseInt(s_page);
		int perPageRow = Integer.parseInt(s_rows);

		int total = 0;
		int records = 0;
		int start = 0;
		int end = 0;

		end = perPageRow * page;
		start = end - (perPageRow - 1);

		commandMap.put("CODE", code);
		commandMap.put("START", start);
		commandMap.put("END", end);

		List<Map<String, Object>> list = sampleService.selectAdminCode(commandMap.getMap());

		records = (int) list.get(0).get("TOTAL_COUNT");

		if (records > 0)
			total = (int) Math.ceil((double) records / (double) perPageRow);
		else
			total = 0;

		mv.addObject("list", list);
		mv.addObject("page", page);
		mv.addObject("total", total);
		mv.addObject("records", records);

		return mv;
	}

	@RequestMapping(value = "/admin/editUserGridView.do")
	public @ResponseBody void editUserJqgrid(@RequestBody String s_data, CommandMap commandMap) throws Exception {
		{
			JSONParser jsonParser = new JSONParser();
			ObjectMapper mapper = new ObjectMapper();
			Map<String, Object> map = new HashMap<String, Object>();

			String data = URLDecoder.decode(s_data, "utf-8").replace("=", "");
			String jsn, userid, Gubun;

			// JSON데이터를 넣어 JSON Object 로 만들어 준다.
			JSONObject jsonObject = (JSONObject) jsonParser.parse(data);

			jsn = jsonObject.get("s_data").toString();
			userid = jsonObject.get("userid").toString();

			map = mapper.readValue(jsn, new TypeReference<HashMap<String, Object>>() {
			});
			map.put("userid", userid);
			Gubun = map.get("oper").toString();

			System.out.println(Gubun);

			switch (Gubun.toUpperCase()) {
			case "EDIT":
				sampleService.updateUserGrid(map);
				break;

			case "ADD":
				sampleService.addUserGrid(map);
				break;

			case "DEL":
				sampleService.delUserGrid(map);
				break;
			}
		}
	}

	// {"s_data":{"CODE_NM":"비밀유지","CODE_SNM":"","HIGH_CODE":"A","oper":"edit","id":"A02"}}=
	@RequestMapping(value = "/admin/editGridView.do")
	public @ResponseBody void editJqgrid(@RequestBody String s_data, CommandMap commandMap) throws Exception {

		JSONParser jsonParser = new JSONParser();
		ObjectMapper mapper = new ObjectMapper();
		Map<String, Object> map = new HashMap<String, Object>();

		String data = URLDecoder.decode(s_data, "utf-8").replace("=", "");

		// JSON데이터를 넣어 JSON Object 로 만들어 준다.
		JSONObject jsonObject = (JSONObject) jsonParser.parse(data);

		String jsn = jsonObject.get("s_data").toString();
		String userid = jsonObject.get("userid").toString();

		map = mapper.readValue(jsn, new TypeReference<HashMap<String, Object>>() {
		});
		map.put("userid", userid);

		System.out.println(map.get("oper"));

		if (map.get("oper").equals("edit")) {
			sampleService.updateGrid(map);
		} else if (map.get("oper").equals("add")) {
			sampleService.addGrid(map);
		} else if (map.get("oper").equals("del")) {
			sampleService.delGrid(map);
		}

	}

	@RequestMapping(value = "/Login/LoginHistory.do", method = RequestMethod.GET)
	public ModelAndView SetLoginHistory(CommandMap commandMap
										, @RequestParam(value = "TYPE", required = false) String type) throws Exception {
		
		HttpServletRequest req = ((ServletRequestAttributes)RequestContextHolder.currentRequestAttributes()).getRequest();	//it is for Client Ip address of login user
		User user = (User)SecurityContextHolder.getContext().getAuthentication().getPrincipal();		//it is Class for spring Security of login user
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardList.do");
		
		String clientUserIp = req.getHeader("X-FORWARDED-FOR");
		clientUserId = user.getUsername();
		
	    if (clientUserIp == null) clientUserIp = req.getRemoteAddr();
	    if (type == null || type.length() < 1) type = "UNUSUAL";
	    if (clientUserId == null || clientUserId.length() < 1) clientUserId = "UNUSUAL";
	    
	    commandMap.put("USER_IP", clientUserIp);
		commandMap.put("TYPE", type);
		commandMap.put("USER_ID", clientUserId);
		
		sampleService.setLoginHistoryWrite(commandMap.getMap());
		
		return mv;
	}
	
	@RequestMapping(value = "/sample/openBoardList.do")
	public ModelAndView openSampleBoardList(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardList");
		
		List<Map<String, Object>> standardList = sampleService.getStandardCode();
		mv.addObject("standardList", standardList);
		
		return mv;
	}
	
	@RequestMapping(value = "/sample/openUserinfoChange.do")
	public ModelAndView openUserinfoChange(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/login/login_UserInfoChange");
		mv.addObject("change", commandMap.get("change"));
		
		return mv;
	}
	
	@RequestMapping(value = "/admin/initUserInfo.do")
	public ModelAndView initUserInfo(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView(); 
		Boolean returnValue = false; //패스워드의 변경상태를 저장
		
		//구분자를 확인하여 패스워드를 초기화 할 것인지 업데이트를 할 것인지를 판단
		String type = commandMap.get("type").toString();

		switch(type.toUpperCase()){
			case "I" : //관리자가 사용자의 패스워드를 초기화할 때
				sampleService.initUserInfo(commandMap.getMap());
				mv.setViewName("redirect:/admin/openAdmin_account.do"); //관리자페이지로 이동
				break;
				
			case "U" :	//사용자가 자신의 비밀번호를 변경할 때
				returnValue = sampleService.updateUserInfo(commandMap.getMap());
				
				if(returnValue == true)
					mv.setViewName("redirect:/sample/openUserinfoChange.do?change=true"); //로그인페이지로 이동
				else
					mv.setViewName("redirect:/sample/openUserinfoChange.do?change=false"); //로그인페이지로 이동
				
				break;
		}
		
		return mv;
	}
	
	@RequestMapping(value = "/sample/BoardHstDelete")
	public ModelAndView openSampleBoardHstDelete(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardDetail.do");
		
		//삭제 로직
		sampleService.deleteBoardHst(commandMap.getMap());
		
		//삭제 후 이전 체결계약서 화면으로 되돌리기 위한 로직
		mv.addObject("CONTRACT_NO", commandMap.get("CONTRACT_NO"));
		mv.addObject("HST_SEQ", commandMap.get("HST_SEQ"));
		return mv;
	}
	
	
	@RequestMapping(value = "/sample/updateBoarHst.do")
	public ModelAndView updateBoardHst(CommandMap commandMap, HttpServletRequest request) throws Exception{
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardHst.do");
		
		//업데이트 로직
		sampleService.updateBoardHst(commandMap.getMap(), request);
		
		//업데이트 후 이전 체결계약서 화면으로 되돌리기 위한 로직
		mv.addObject("CONTRACT_NO", commandMap.get("CONTRACT_NO"));
		mv.addObject("HST_SEQ", commandMap.get("HST_SEQ"));
		
		return mv;
	}
	
	@RequestMapping(value = "/sample/selectBoardList.do")
	public ModelAndView selectBoardList(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("jsonView");

		List<Map<String, Object>> list = sampleService.selectBoardList(commandMap.getMap());
		mv.addObject("list", list);
		
		if (list.size() > 0) {
			mv.addObject("TOTAL", list.get(0).get("TOTAL_COUNT"));
		} else {
			mv.addObject("TOTAL", 0);
		}

		return mv;
	}

	@RequestMapping(value = "/sample/selectBoardSearch.do")
	public ModelAndView selectBoardSearchList(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("jsonView");

		List<Map<String, Object>> list = sampleService.selectBoardSearchList(commandMap.getMap());
		mv.addObject("list", list);

		if (list.size() > 0) {
			mv.addObject("TOTAL", list.get(0).get("TOTAL_COUNT"));
		} else {
			mv.addObject("TOTAL", 0);
		}

		return mv;
	}

	@RequestMapping(value = "/sample/openBoardDetail.do")
	public ModelAndView openBoardDetail(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardDetail");
		//ModelAndView mv = new ModelAndView("/sample/boardContractSelect");

		System.out.println(commandMap.getMap());

		Map<String, Object> map = sampleService.selectBoardDetail(commandMap.getMap());
		mv.addObject("map", map.get("map"));
		mv.addObject("list", map.get("list"));
		mv.addObject("h_list", map.get("h_list"));

		return mv;
	}

	@RequestMapping(value = "/sample/testMapArgumentResolver.do")
	public ModelAndView testMapArgumentResolver(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("");

		if (commandMap.isEmpty() == false) {
			Iterator<Entry<String, Object>> iterator = commandMap.getMap().entrySet().iterator();
			Entry<String, Object> entry = null;
			while (iterator.hasNext()) {
				entry = iterator.next();
				log.debug("key : " + entry.getKey() + ", value : " + entry.getValue());
			}
		}
		return mv;
	}

	@RequestMapping(value = "/sample/openinitBox.do")
	public ModelAndView openinitBox(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("jsonView");

		List<Map<String, Object>> list = sampleService.selectBoxList(commandMap.getMap());
		mv.addObject("list", list);

		return mv;
	}

	@RequestMapping(value = "/sample/openBoardWrite.do")
	public ModelAndView openBoardWrite(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardWrite");

		return mv;
	}

	@RequestMapping(value = "/sample/openBoardUpdate.do")
	public ModelAndView openBoardUpdate(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardUpdate");

		Map<String, Object> map = sampleService.selectBoardUpdateDetail(commandMap.getMap());
		mv.addObject("map", map.get("map"));
		mv.addObject("list", map.get("list"));

		return mv;
	}
	
	@RequestMapping(value = "/sample/openBoardHstUpdate.do")
	public ModelAndView openBoardHstUpdate(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardHstUpdate");

		Map<String, Object> map = sampleService.selectBoardUpdateHstDetail(commandMap.getMap());
		mv.addObject("map", map.get("map"));
		mv.addObject("list", map.get("list"));

		return mv;
	}

	@RequestMapping(value = "/sample/openBoardDelete.do")
	public ModelAndView openBoardDelete(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardList.do");

		sampleService.selectBoardUpdateDetail(commandMap.getMap());

		return mv;
	}
	
	@RequestMapping(value = "/sample/deleteBoard.do")
	public ModelAndView openContractDelete(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardList");
		
		sampleService.deleteContract(commandMap.getMap());

		return mv;
	}

	@RequestMapping(value = "/sample/updateBoard.do")
	public ModelAndView updateBoard(CommandMap commandMap,
									@RequestParam(value = "userid", required = false) String userid,
									HttpServletRequest request
									) throws Exception {
		
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardDetail.do");

		commandMap.put("userid", userid);
		sampleService.updateBoard(commandMap.getMap(), request);

		mv.addObject("CONTRACT_NO", commandMap.get("CONTRACT_NO"));
		return mv;
	}

	@RequestMapping(value = "/sample/openBoardModify.do")
	public ModelAndView openBoardModify(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardModify");

		return mv;
	}

	@RequestMapping(value = "/sample/insertBoard.do")
	public ModelAndView insertBoard(@RequestParam(value = "userid", required = false) String userid,
									CommandMap commandMap, 
									HttpServletRequest request
									) throws Exception {
		
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardList.do");
		commandMap.put("userid", userid);
		
		sampleService.insertBoard(commandMap.getMap(), request);

		return mv;
	}

	@RequestMapping(value = "/sample/modifyBoard.do")
	public ModelAndView modifyBoard(CommandMap commandMap, HttpServletRequest request) throws Exception {
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardList.do");
		
		sampleService.modifyBoard(commandMap.getMap(), request);

		return mv;
	}

	@RequestMapping(value = "/sample/endBoard.do")
	public ModelAndView endBoard(CommandMap commandMap, HttpServletRequest request) throws Exception {
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardList.do");

		sampleService.endBoard(commandMap.getMap(), request);

		return mv;
	}

	@RequestMapping(value = "/sample/openPopup.do")
	public ModelAndView openPopup(@RequestParam Map<String, Object> commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardPopup");

		List<Map<String, Object>> list = sampleService.openPopupList(commandMap);
		mv.addObject("list", list);

		return mv;
	}

	@RequestMapping(value = "/sample/openContractPopup.do", method={RequestMethod.GET, RequestMethod.POST})
	public ModelAndView openContractPopup(@RequestParam Map<String, Object> commandMap, HttpServletRequest request) throws Exception {
		String userid = request.getParameter("userid");
		
		ModelAndView mv = new ModelAndView("/sample/boardContractPopup");
		commandMap.put("userid", userid);
		
		List<Map<String, Object>> list = sampleService.openContractPopupList(commandMap);
		mv.addObject("list", list);

		return mv;
	}

	@RequestMapping(value = "/sample/openGubunPopup.do")
	public ModelAndView openGubunPopup(@RequestParam Map<String, Object> commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardGubunPopup");

		String id = commandMap.get("id").toString();
		
		switch(id){
			case "BU_NM" :
				commandMap.put("GUBUNTYPE", "B"); //기준정보 B%를  찾기 위함(사업부)
				break;

			default :
				commandMap.put("GUBUNTYPE", "A"); //기준정보 A%를  찾기 위함(계약구분)
				break;
		}
		
		List<Map<String, Object>> list = sampleService.openGubunPopupList(commandMap);
		mv.addObject("list", list);

		return mv;
	}

	@RequestMapping(value = "/sample/openBoardHst.do")
	public ModelAndView openBoardHst(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("/sample/boardHst");

		Map<String, Object> map = sampleService.openHstList(commandMap.getMap());
		mv.addObject("map", map.get("map"));
		mv.addObject("list", map.get("list"));

		return mv;
	}

	@RequestMapping(value = "/sample/openUpdateContent.do", method={RequestMethod.GET, RequestMethod.POST})
	public ModelAndView openUpdateContent(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardDetail.do");

		sampleService.updateCotents(commandMap.getMap());

		mv.addObject("CONTRACT_NO", commandMap.get("CONTRACT_NO"));

		return mv;
	}

	@RequestMapping(value = "/sample/openSendEmail.do")
	public ModelAndView openSendEmail(CommandMap commandMap) throws Exception {
		ModelAndView mv = new ModelAndView("redirect:/sample/openBoardList.do");

		sampleService.sendEmail(commandMap.getMap());

		return mv;
	}

	@RequestMapping(value = "/report/openReportView.do")
	public ModelAndView openReportView( @RequestParam(value = "val", required = false) String gubun,
										@RequestParam(value = "year", required = false) String year, 
										@RequestParam(value = "userid", required = false) String userid,
										@RequestParam(value = "G_gubun", required = false) String G_gubun, 
										CommandMap commandMap
									  ) throws Exception {
		ModelAndView mv = new ModelAndView("jsonView");

		commandMap.put("year", year);
		commandMap.put("userid", userid);

		if (gubun.equals("A")) {
			Map<String, Object> map = sampleService.NtotalContract(commandMap.getMap());

			mv.addObject("graph", map.get("graph"));
			mv.addObject("grid", map.get("grid"));
		} else if (gubun.equals("B")) {
			List<Map<String, Object>> list = sampleService.NperiodContract(commandMap.getMap());

			mv.addObject("graph", list);
		} else if (gubun.equals("C")) {
			List<Map<String, Object>> list;

			switch (G_gubun) {
			case "simple":
				list = sampleService.AtotalContract_Simple();
				break;

			case "detail":
				list = sampleService.AtotalContract_Detail();

			default: // 예외를 대비하여 설정, 만약 아무것도 아닐 경우 소분류로 구분하여 화면에 출력
				list = sampleService.AtotalContract_Detail();
			}

			mv.addObject("grid", list);
		} else if (gubun.equals("D")) {
			List<Map<String, Object>> list = sampleService.AperiodContract(commandMap.getMap());

			mv.addObject("grid", list);
		}

		return mv;
	}

}
