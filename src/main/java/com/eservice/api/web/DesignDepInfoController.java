package com.eservice.api.web;
import com.alibaba.fastjson.JSON;
import com.eservice.api.core.Result;
import com.eservice.api.core.ResultGenerator;
import com.eservice.api.model.design_dep_info.DesignDepInfo;
import com.eservice.api.model.design_dep_info.DesignDepInfoDetail;
import com.eservice.api.service.common.CommonService;
import com.eservice.api.service.common.Constant;
import com.eservice.api.service.impl.DesignDepInfoServiceImpl;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.interceptor.TransactionAspectSupport;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import java.io.File;
import java.util.Date;
import java.util.List;

/**
* Class Description: xxx
* @author Wilson Hu
* @date 2020/06/07.
*/
@RestController
@RequestMapping("/design/dep/info")
public class DesignDepInfoController {
    @Resource
    private DesignDepInfoServiceImpl designDepInfoService;

    @Value("${design_attached_saved_dir}")
    private String designAttachedSavedDir;

    @Resource
    private CommonService commonService;

    private Logger logger = Logger.getLogger(DesignDepInfoController.class);
    /**
     * 一次性同时上传 信息
     *
     */
    @Transactional(rollbackFor = Exception.class)
    @PostMapping("/add")
    public Result add(String jsonDesignDepInfoFormAllInfo) {
        DesignDepInfo designDepInfo = JSON.parseObject(jsonDesignDepInfoFormAllInfo, DesignDepInfo.class);
        if (designDepInfo == null || designDepInfo.equals("")) {
            return ResultGenerator.genFailResult("JSON数据异常");
        }
        if(designDepInfo.getOrderId() == null){
            return ResultGenerator.genFailResult("异常，orderId为空");
        }
        //自动设置状态
        if(designDepInfo.getDesigner() == null || designDepInfo.getDesigner().equals("")){
            designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_UNPLANNED);
        } else {
            designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_PLANNED);
        }
        designDepInfoService.saveAndGetID(designDepInfo);
        // 返回ID给前端，前端新增联系单时不关闭页面。
        return ResultGenerator.genSuccessResult(designDepInfo.getId());
    }

    @PostMapping("/delete")
    public Result delete(@RequestParam Integer id) {
        designDepInfoService.deleteById(id);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/update")
    public Result update(String jsonDesignDepInfoFormAllInfo) {
        DesignDepInfo designDepInfo = JSON.parseObject(jsonDesignDepInfoFormAllInfo, DesignDepInfo.class);

        if (designDepInfo == null || designDepInfo.equals("")) {
            return ResultGenerator.genFailResult("JSON数据异常");
        }
        if(designDepInfo.getId() == null){
            return ResultGenerator.genFailResult("异常，Id为空");
        }
        //自动设置状态
        if(designDepInfo.getDesigner() == null || designDepInfo.getDesigner().equals("")){
            designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_UNPLANNED);
        } else {
            designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_PLANNED);
        }

        DesignDepInfo designDepInfoOld = designDepInfoService.findById(designDepInfo.getId());
        if(designDepInfoOld == null){
            return ResultGenerator.genFailResult("异常，根据Id找不到对应的设计");
        }
        //记录的是: 最新一次更新
        //整体状态DesignStatus, 只考虑常规操作, 即只考虑从未完成到已完成, 不考虑来回改动,
        if(designDepInfoOld.getDrawingFileDone() != designDepInfo.getDrawingFileDone()){
            if(designDepInfo.getDrawingFileDone() ) {
                designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_DRAWING_DONE);
            } else {
                //对于把已某项 完成改成未完成的情况,不考虑再修改整体状态
            }
        }
        if(designDepInfoOld.getLoadingFileDone() != designDepInfo.getLoadingFileDone()){
            if(designDepInfo.getLoadingFileDone() ) {
                designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_LOADING_DONE);
            }
        }
        if(designDepInfoOld.getHoleDone() != designDepInfo.getHoleDone()){
            if(designDepInfo.getHoleDone() ) {
                designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_HOLE_DONE);
            }
        }
        if(designDepInfoOld.getTubeDone() != designDepInfo.getTubeDone()){
            if(designDepInfo.getTubeDone() ) {
                designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_TUBE_DONE);
            }
        }
        if(designDepInfoOld.getCoverDone() != designDepInfo.getCoverDone()){
            if(designDepInfo.getCoverDone() ) {
                designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_COVER_DONE);
            }
        }
        if(designDepInfoOld.getBomDone() != designDepInfo.getBomDone()){
            if(designDepInfo.getBomDone() ) {
                designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_BOM_DONE);
            }
        }
        if(designDepInfo.getDrawingFileDone()
                &&designDepInfo.getLoadingFileDone()
                &&designDepInfo.getHoleDone()
                &&designDepInfo.getTubeDone()
                &&designDepInfo.getCoverDone()
                &&designDepInfo.getBomDone() ){
            designDepInfo.setDesignStatus(Constant.STR_DESIGN_STATUS_ALL_DONE);
        }

        designDepInfoService.update(designDepInfo);
        return ResultGenerator.genSuccessResult();
    }

    /**
     * 后端提示：No mapping found for HTTP request with URI [/design/depinfo/uploadDesignFiles]
     * in DispatcherServlet with name 'dispatcherServlet.
     * 其实是接口名称没有完全正确，多了个字符。
     */
    @PostMapping("/uploadDesignFile")
    public Result uploadDesignFile(HttpServletRequest request) {
        try {
            String orderNum = request.getParameterValues("orderNum")[0];
            String type = request.getParameterValues("type")[0];
            List<MultipartFile> fileList = ((MultipartHttpServletRequest) request).getFiles("file");
            if (fileList == null || fileList.size() == 0) {
                return ResultGenerator.genFailResult("Error: no available file");
            }
            MultipartFile file = fileList.get(0);
            File dir = new File(designAttachedSavedDir);
            if (!dir.exists()) {
                dir.mkdir();
            }

            try {
                // save file， 只保存文件，不保存数据库，保存路径返回给前端，前端统一写入 。
                String resultPath = commonService.saveFile(designAttachedSavedDir, file, type, orderNum,
                        Constant.DESIGN_ATTACHED_FILE, 0);
                if (resultPath == null || resultPath == "") {
                    return ResultGenerator.genFailResult("文件名为空，设计附件上传失败！");
                }
                logger.info("====/design/dep/info uploadDesignFile(): success======== " + resultPath);
                /**
                 * 在第一次新建时上传文件，designDepInfoID为空，则后端把各个更新日期设为当前日期即可
                 * 在编辑时更新上传文件，则把根据类型，把对应的Update日期更新。
                 */
//                if(request.getParameterValues("designDepInfoID") == null){
//                    ///todo ,第一次新建时上传文件，会出错
//                }
//                Integer designDepInfoID = Integer.valueOf(request.getParameterValues("designDepInfoID")[0]);
//                if(designDepInfoID == null || designDepInfoID.equals(0)){
//                    logger.info("在新创建的设计里上传文件，已经默认更新日期为当前时间");
//                } else {
//                    DesignDepInfo designDepInfo = designDepInfoService.findById(designDepInfoID);
//                    if(designDepInfo == null){
//                        logger.error("根据前端给的ID找不到 designDepInfo");
//                    } else {
//                        logger.info("type文件类型：" + type);
//                        switch (type) {
//                            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_DRAWING:
//                                designDepInfo.setDrawingUpdateTime(new Date());
//                                break;
//                            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_LOADINGFILE:
//                                designDepInfo.setLoadingUpdateTime(new Date());
//                                break;
//                            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_HOLE:
//                                designDepInfo.setHoleUpdateTime(new Date());
//                                break;
//                            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_TUBE:
//                                designDepInfo.setTubeUpdateTime(new Date());
//                                break;
//                            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_BOM:
//                                designDepInfo.setBomUpdateTime(new Date());
//                                break;
//                            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_COVER:
//                                designDepInfo.setCoverUpdateTime(new Date());
//                                break;
//
//                            default:
//                                break;
//                        }
//                        designDepInfoService.update(designDepInfo);
//                    }
//                }
                return ResultGenerator.genSuccessResult(resultPath);
            } catch (Exception e) {
                e.printStackTrace();
                TransactionAspectSupport.currentTransactionStatus().setRollbackOnly();
                return ResultGenerator.genFailResult("设计附件 上传失败！" + e.getMessage());
            }

        } catch (Exception e) {
            e.printStackTrace();
            return ResultGenerator.genFailResult("设计附件 上传失败！" + e.getMessage());
        }
    }

    @PostMapping("/detail")
    public Result detail(@RequestParam Integer id) {
        DesignDepInfo designDepInfo = designDepInfoService.findById(id);
        return ResultGenerator.genSuccessResult(designDepInfo);
    }

    @PostMapping("/list")
    public Result list(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size) {
        PageHelper.startPage(page, size);
        List<DesignDepInfo> list = designDepInfoService.findAll();
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }

    @PostMapping("/selectDesignDepInfo")
    public Result selectDesignDepInfo(@RequestParam(defaultValue = "0") Integer page,
                                      @RequestParam(defaultValue = "0") Integer size,
                                      String orderNum,
                                      String saleman,
                                      String guestName,
                                      Integer orderStatus,//订单状态
                                      Integer drawingLoadingDoneStatus,//图纸状态
                                      String machineSpec,
                                      String keywords,
                                      String designer,
                                      String updateDateStart,
                                      String updateDateEnd) {
        PageHelper.startPage(page, size);
        List<DesignDepInfoDetail> list = designDepInfoService.selectDesignDepInfo(
                orderNum,
                saleman,
                guestName,
                orderStatus,
                drawingLoadingDoneStatus,
                machineSpec,
                keywords,
                designer,
                updateDateStart,
                updateDateEnd);
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }

    /**
     * 根据 设计单的ID以及下载类型（） 返回设计单的附件的对应类型的文件名称
     * 下载路径的前面部分是统一的，放在xxx_ip/download/下(nginx配置)，
     * 比如，访问下面地址可以下载该图纸装车单
     *
     * @param designDepInfoId
     * @return 类似
     */
    @PostMapping("/getDesignAttachedFile")
    public Result getDesignAttachedFile(@RequestParam Integer designDepInfoId, String fileType) {

        DesignDepInfo ddi = designDepInfoService.findById(designDepInfoId);
        if (null == ddi) {
            return ResultGenerator.genFailResult("根据该 designDepInfoId 找不到对应的设计单");
        }

        String fileName = null;
        switch (fileType){
            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_DRAWING:
                if (ddi.getDrawingFiles() == null) {
                    return ResultGenerator.genFailResult("该设计单没有 图纸附件");
                }
                fileName = ddi.getDrawingFiles().substring(designAttachedSavedDir.length());
                break;

            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_LOADINGFILE:
                if (ddi.getLoadingFiles() == null) {
                    return ResultGenerator.genFailResult("该设计单没有 装车单附件");
                }
                fileName = ddi.getLoadingFiles().substring(designAttachedSavedDir.length());
                break;

            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_HOLE:
                if (ddi.getHoleFiles() == null) {
                    return ResultGenerator.genFailResult("该设计单没有 点孔附件");
                }
                fileName = ddi.getHoleFiles().substring(designAttachedSavedDir.length());
                break;
            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_TUBE:
                if (ddi.getTubeFiles() == null) {
                    return ResultGenerator.genFailResult("该设计单没有 方管附件");
                }
                fileName = ddi.getTubeFiles().substring(designAttachedSavedDir.length());
                break;

            case Constant.STR_DESIGN_UPLOAD_FILE_TYPE_COVER:
                if (ddi.getCoverFile() == null) {
                    return ResultGenerator.genFailResult("该设计单没有 罩盖附件");
                }
                fileName = ddi.getCoverFile().substring(designAttachedSavedDir.length());
                break;

                default:
                    return ResultGenerator.genFailResult("没有该类型： " + fileType);
        }

        return ResultGenerator.genSuccessResult(fileName);
    }
}
