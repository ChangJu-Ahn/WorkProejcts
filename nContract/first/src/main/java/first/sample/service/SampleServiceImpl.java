package first.sample.service;

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;

import org.apache.log4j.Logger;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import first.common.util.FileUtils;
import first.sample.dao.SampleDAO;

@Service("sampleService")
public class SampleServiceImpl implements SampleService {
	Logger log = Logger.getLogger(this.getClass());

	@Resource(name = "fileUtils")
	private FileUtils fileUtils;

	@Resource(name = "sampleDAO")
	private SampleDAO sampleDAO;

	@Override
	public List<Map<String, Object>> selectBoardList(Map<String, Object> map) throws Exception {
		String tempBusorValue = map.get("BUSOR_CODE").toString().trim();
		
		if(tempBusorValue.equals("ALL"))
			return sampleDAO.selectBoardList(map);
		else
			return sampleDAO.selectBoardBusorList(map);
	}

	@Override
	public List<Map<String, Object>> openPopupList(Map<String, Object> map) throws Exception {
		return sampleDAO.openPopupList(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void insertBoard(Map<String, Object> map, HttpServletRequest request) throws Exception {

		sampleDAO.insertBoard(map);				//한 줄은 이전 시퀀스(나중에 다 완료되면 지울 것)
		sampleDAO.insertHeaderBoard(map);		//한 줄은 변경 후 시퀀스(나중에 다 완료되면 사용할 것)		
		
		List<Map<String, Object>> list = setMultipartHttpServletRequest(map, request);
		
		for (int i = 0, size = list.size(); i < size; i++) {
			sampleDAO.insertFile(list.get(i));
		}
	
		//아래 두 줄은 이전 시퀀스(파일 시퀀스 업로드 부분)
		sampleDAO.updateFileSeq(map);		   
		sampleDAO.insertHst(map);
		
		//아래 세 줄은 변경 시퀀스
		sampleDAO.updateFileSeqToHeader(map);
		sampleDAO.insertDetailBoard(map);
		sampleDAO.insertHistoryBoard(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void modifyBoard(Map<String, Object> map, HttpServletRequest request) throws Exception {
		
		sampleDAO.modifyBoard(map);
		
		List<Map<String, Object>> list = fileUtils.parseInsertFileInfo(map, request);
		for (int i = 0, size = list.size(); i < size; i++) {
			sampleDAO.modifyFile(list.get(i));
		}
		
		if(list.size() > 0){
			sampleDAO.updateFileSeq(map);
		}
		
		sampleDAO.insertHst(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void endBoard(Map<String, Object> map, HttpServletRequest request) throws Exception{
		
		sampleDAO.endBoard(map);
		
		List<Map<String, Object>> list = fileUtils.parseInsertFileInfo(map, request);
		for (int i = 0, size = list.size(); i < size; i++) {
			sampleDAO.modifyFile(list.get(i));
		}
		
		if(list.size() > 0){
			sampleDAO.updateFileSeq(map);
		}
		
		sampleDAO.insertHst(map);
	}
	
	@Override
	public List<Map<String, Object>> openContractPopupList(Map<String, Object> map) throws Exception  {
		return sampleDAO.openContractPopupList(map);
	}

	@Override
	public Map<String, Object> selectBoardDetail(Map<String, Object> map) throws Exception{
		
		Map<String, Object> resultMap = new HashMap<String,Object>();
		
		Map<String,Object> t_map = sampleDAO.selectBoardDetail(map);
		List<Map<String, Object>> list = sampleDAO.selectFileList(map);
		List<Map<String, Object>> h_list = sampleDAO.selectHstList(map);
		
		resultMap.put("map", t_map);
		resultMap.put("list", list);
		resultMap.put("h_list", h_list);
		
		return resultMap;
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void updateBoard(Map<String, Object> map, HttpServletRequest request) throws Exception {
		Map<String,Object> tempMap = null;
		
		sampleDAO.updateBoard(map);
	 	sampleDAO.deleteFileList(map);
	 
	    List<Map<String,Object>> list = setMultipartHttpServletRequest(map, request);
	    
	    for(int i=0, size=list.size(); i<size; i++){
	        tempMap = list.get(i);
	        
	        if(tempMap.get("IS_NEW").equals("Y")){
	            sampleDAO.insertFile(tempMap);
	        }
	        else{
	            sampleDAO.updateFile(tempMap);
	        }
	    }
	    //재고 ~~~~~
	    //sampleDAO.insertHst(map);
	}

	@Override
	public Map<String, Object> selectBoardUpdateDetail(Map<String, Object> map) {
		Map<String, Object> resultMap = new HashMap<String,Object>();
		
		Map<String,Object> t_map = sampleDAO.selectUpdateBoardDetail(map);
		List<Map<String, Object>> list = sampleDAO.selectFileList(map);
		
		resultMap.put("map", t_map);
		resultMap.put("list", list);
		
		return resultMap;
	}

	@Override
	public Map<String, Object> openHstList(Map<String, Object> map) {
		Map<String, Object> resultMap = new HashMap<String,Object>();
		
		Map<String,Object> t_map = sampleDAO.selectBoardHstDetail(map);
		List<Map<String, Object>> list = sampleDAO.selectHstFileList(map);
		
		resultMap.put("map", t_map);
		resultMap.put("list", list);
		
		return resultMap;
	}

	@Override
	public List<Map<String, Object>> openGubunPopupList(Map<String, Object> map) {
		return sampleDAO.openGubunPopupList(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void updateCotents(Map<String, Object> map) {
		sampleDAO.updateContents(map);
	}

	@Override
	public void sendEmail(Map<String, Object> map) {
		sampleDAO.sendEmail(map);
	}

	@Override
	public List<Map<String, Object>> selectBoardSearchList(Map<String, Object> map) throws Exception{
		String tempBusorValue = map.get("BUSOR_CODE").toString().trim();
		
		if(tempBusorValue.equals("ALL"))
			return sampleDAO.selectBoardSearchList(map);
		else
			return sampleDAO.selectBoardBusorSearchList(map);
	}

	@Override
	public List<Map<String, Object>> selectAdminCode(Map<String, Object> map) throws Exception {
		return sampleDAO.selectBoardAdminCode(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void updateGrid(Map<String, Object> map) throws Exception {
		sampleDAO.updateGrid(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void addGrid(Map<String, Object> map) throws Exception {
		sampleDAO.addGrid(map);
		
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void delGrid(Map<String, Object> map) throws Exception {
		sampleDAO.delGrid(map);
	}

	@Override
	public Map<String, Object> NtotalContract(Map<String, Object> map) throws Exception {
		Map<String, Object> resultMap = new HashMap<String,Object>();
		
		resultMap.put("graph",sampleDAO.NtotalContractGraph(map));
		resultMap.put("grid",sampleDAO.NtotalContractGrid(map));
		
		return resultMap;
	}

	@Override
	public List<Map<String, Object>> NperiodContract(Map<String, Object> map) throws Exception {
		return sampleDAO.NperiodContract(map);
	}

	@Override
	public List<Map<String, Object>> AperiodContract(Map<String, Object> map) throws Exception {
		return sampleDAO.AperiodContract(map);
	}

	@Override
	public List<Map<String, Object>> AtotalContract_Detail() throws Exception {
		return sampleDAO.AtotalContract_Detail();
	}

	@Override
	public List<Map<String, Object>> AtotalContract_Simple() throws Exception {
		return sampleDAO.AtotalContract_Simple();
	}
	
	@Override
	public List<Map<String, Object>> selectBoxList(Map<String, Object> map) throws Exception {
		return sampleDAO.selectBoxList(map);
	}

	@Override
	public List<Map<String, Object>> getAllData(Map<String, Object> map) throws Exception { 
		String tempTarget = map.get("TARGET").toString().toUpperCase();
		List<Map<String, Object>> tempList = null;
		
		//execute DAO Method by Target Value.
		switch(tempTarget.toUpperCase()){
			 case "CONTRACT_LIST" :
				 tempList = sampleDAO.getAllData(map);
				 break;
				 
			 case "ADMINCODE_LIST" :
				 tempList = sampleDAO.getAllAdminCodeList(map);
				 break;
		}
		
		return tempList;
	}

	@Override
	public List<Map<String, Object>> selectUserList(Map<String, Object> map) throws Exception {
		return sampleDAO.selectUserList(map);
	}
	
	@Override
	public List<Map<String, Object>> getStandardCode() throws Exception {
		return sampleDAO.getStandardCode();
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void updateUserGrid(Map<String, Object> map) throws Exception {
		sampleDAO.updateUserGrid(map);
	}
	
	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void addUserGrid(Map<String, Object> map) throws Exception {
		sampleDAO.addUserGrid(map);
	}
	
	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void delUserGrid(Map<String, Object> map) throws Exception {
		sampleDAO.delUserGrid(map);
	}
	
	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void deleteContract(Map<String, Object> map) throws Exception {
		sampleDAO.deleteContract(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void deleteBoardHst(Map<String, Object> map) throws Exception {
		sampleDAO.deleteContractHst(map);
	}
	
	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void updateBoardHst(Map<String, Object> map, HttpServletRequest request) throws Exception {
		sampleDAO.updateContractHst(map);
		
		MultipartHttpServletRequest multipartHttpServletRequest = (MultipartHttpServletRequest) request;
		Iterator<String> iterator = multipartHttpServletRequest.getFileNames();
		MultipartFile multipartFile = null;
		while (iterator.hasNext()) {
			multipartFile = multipartHttpServletRequest.getFile(iterator.next());
			if (multipartFile.isEmpty() == false) {
				log.debug("------------- file start -------------");
				log.debug("name : " + multipartFile.getName());
				log.debug("filename : " + multipartFile.getOriginalFilename());
				log.debug("size : " + multipartFile.getSize());
				log.debug("-------------- file end --------------\n");
			}
		}
		
	 	sampleDAO.deleteFileList(map);
	 
	    List<Map<String,Object>> list = fileUtils.parseUpdateFileInfo(map, request);
	    Map<String,Object> tempMap = null;
	    for(int i=0, size=list.size(); i<size; i++){
	        tempMap = list.get(i);
	        
	        if(tempMap.get("IS_NEW").equals("Y")){
	            sampleDAO.insertHstFile(tempMap);
	        }
	        else{
	            sampleDAO.updateFile(tempMap);
	        }
	    }
		
	}
	
	@Override
	public Map<String, Object> selectBoardUpdateHstDetail(Map<String, Object> map) {
		Map<String, Object> resultMap = new HashMap<String,Object>();
		
		Map<String,Object> t_map = sampleDAO.selectUpdateBoardHstDetail(map);
		List<Map<String, Object>> list = sampleDAO.selectHstFileList(map);
		
		resultMap.put("map", t_map);
		resultMap.put("list", list);
		
		return resultMap;
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void initUserInfo(Map<String, Object> map) {
		sampleDAO.updateUserInfo_initial(map);
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public boolean updateUserInfo(Map<String, Object> map) {
		return sampleDAO.updateUserInfo_update(map); 		//변경에 성공하면 true, 실패하면 false을 반환하여 컨트롤 단에서 페이지를 구분
	}

	@Override
	@Transactional(propagation = Propagation.REQUIRED, rollbackFor={Exception.class})
	public void setLoginHistoryWrite(Map<String, Object> map) {
		String type = map.get("TYPE").toString().toUpperCase();
		
		if(type.equals("LOGIN")) 
			sampleDAO.insertLoginHistory(map);
		else 
			sampleDAO.updateLoginHistory(map);
	}
	
	public List<Map<String, Object>> setMultipartHttpServletRequest(Map<String, Object> map, HttpServletRequest request) throws Exception {
		MultipartHttpServletRequest multipartHttpServletRequest = (MultipartHttpServletRequest) request;
		Iterator<String> iterator = multipartHttpServletRequest.getFileNames();
		MultipartFile multipartFile = null;
		
		while (iterator.hasNext()) {
			multipartFile = multipartHttpServletRequest.getFile(iterator.next());
			if (multipartFile.isEmpty() == false) {
				log.debug("------------- file start -------------");
				log.debug("name : " + multipartFile.getName());
				log.debug("filename : " + multipartFile.getOriginalFilename());
				log.debug("size : " + multipartFile.getSize());
				log.debug("-------------- file end --------------\n");
			}
		}

		return fileUtils.parseInsertFileInfo(map, request);
	}
	
}

