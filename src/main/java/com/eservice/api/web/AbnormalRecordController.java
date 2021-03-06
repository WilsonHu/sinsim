package com.eservice.api.web;

import com.alibaba.fastjson.JSON;
import com.eservice.api.core.Result;
import com.eservice.api.core.ResultGenerator;
import com.eservice.api.model.abnormal.Abnormal;
import com.eservice.api.model.abnormal_image.AbnormalImage;
import com.eservice.api.model.abnormal_record.AbnormalRecord;
import com.eservice.api.model.abnormal_record.AbnormalRecordDetail;
import com.eservice.api.model.machine.Machine;
import com.eservice.api.model.machine_order.MachineOrder;
import com.eservice.api.model.process_record.ProcessRecord;
import com.eservice.api.model.task.Task;
import com.eservice.api.model.task_record.TaskRecord;
import com.eservice.api.service.AbnormalImageService;
import com.eservice.api.service.AbnormalService;
import com.eservice.api.service.UserService;
import com.eservice.api.service.common.CommonService;
import com.eservice.api.service.common.Constant;
import com.eservice.api.service.impl.*;
import com.eservice.api.service.mqtt.MqttMessageHelper;
import com.eservice.api.service.mqtt.ServerToClientMsg;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import tk.mybatis.mapper.entity.Condition;

import javax.annotation.Resource;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * Class Description: xxx
 *
 * @author Wilson Hu
 * @date 2017/12/27.
 */
@RestController
@RequestMapping("/abnormal/record")
public class AbnormalRecordController {
    @Resource
    private AbnormalRecordServiceImpl abnormalRecordService;
    @Resource
    private AbnormalService abnormalService;
    @Resource
    private AbnormalImageService abnormalImageService;
    @Resource
    private TaskRecordServiceImpl taskRecordService;
    @Resource
    private MachineServiceImpl machineService;
    @Resource
    private ProcessRecordServiceImpl processRecordService;
    @Resource
    private TaskServiceImpl taskService;
    @Resource
    private MachineOrderServiceImpl machineOrderService;
    @Resource
    private MqttMessageHelper mqttMessageHelper;
    @Resource
    private CommonService commonService;
    /**
     * 异常管理excel表格，和合同excel表格放同个地方
     */
    @Value("${contract_excel_output_dir}")
    private String abnoramlExcelOutputDir;
    @Resource
    private UserService userService;


    @PostMapping("/add")
    public Result add(String abnormalRecord) {
        AbnormalRecord abnormalRecord1 = JSON.parseObject(abnormalRecord, AbnormalRecord.class);

        abnormalRecordService.save(abnormalRecord1);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/delete")
    public Result delete(@RequestParam Integer id) {
        abnormalRecordService.deleteById(id);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/update")
    @Transactional(rollbackFor = Exception.class)
    public Result update(String abnormalRecord) {
        AbnormalRecord abnormalRecord1 = JSON.parseObject(abnormalRecord, AbnormalRecord.class);
        abnormalRecord1.setSolveTime(new Date());
        if(abnormalRecord1.getSolution()==null || abnormalRecord1.getSolution().isEmpty()){
//            abnormalRecord1.setSolution("空");
            return ResultGenerator.genFailResult("请填写解决内容");
        }
        //修改对应工序的状态为"安装中"或者“质检中”，需要检查安装开始时间和质检开始时间，质检已开始则状态为质检中
        AbnormalRecord completeInfo = abnormalRecordService.findById(abnormalRecord1.getId());
        Integer taskRecordId = completeInfo.getTaskRecordId();
        if (taskRecordId != null && taskRecordId > 0) {
            TaskRecord tr = taskRecordService.findById(taskRecordId);
            //MQTT 异常解决后，通知工序的安装组长
            String taskName = tr.getTaskName();
            Condition condition = new Condition(Task.class);
            condition.createCriteria().andCondition("task_name = ", taskName);
            List<Task> taskList = taskService.findByCondition(condition);
            if (taskList == null || taskList.size() <= 0) {
                throw new RuntimeException();
            }
            tr.setStatus(Constant.TASK_INSTALLING);
            taskRecordService.update(tr);
            //更新task record状态时候，必须去更新process record中对应task的状态
            commonService.updateTaskRecordRelatedStatus(tr);
            ProcessRecord pr = processRecordService.findById(tr.getProcessRecordId());
            Machine machine = machineService.findById(pr.getMachineId());
            MachineOrder machineOrder = machineOrderService.findById(machine.getOrderId());
            ServerToClientMsg msg = new ServerToClientMsg();
            msg.setOrderNum(machineOrder.getOrderNum());
            msg.setNameplate(machine.getNameplate());
            mqttMessageHelper.sendToClient(Constant.S2C_INSTALL_ABNORMAL_RESOLVE + taskList.get(0).getGroupId(), JSON.toJSONString(msg));
        } else {
            throw new RuntimeException();
        }
        abnormalRecordService.update(abnormalRecord1);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/detail")
    public Result detail(@RequestParam Integer id) {
        AbnormalRecord abnormalRecord = abnormalRecordService.findById(id);
        return ResultGenerator.genSuccessResult(abnormalRecord);
    }

    @PostMapping("/list")
    public Result list(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size) {
        PageHelper.startPage(page, size);
        List<AbnormalRecord> list = abnormalRecordService.findAll();
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }

    /**
     * 根据 task_record.id 返回abnormalRecordDetail
     *
     * @param taskRecordId
     * @return
     */
    @PostMapping("/selectAbnormalRecordDetails")
    public Result selectAbnormalRecordDetails(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size,
                                              @RequestParam Integer taskRecordId) {
        PageHelper.startPage(page, size);
        List<AbnormalRecordDetail> abnormalRecordDetailList = abnormalRecordService.selectAbnormalRecordDetails(taskRecordId);
        PageInfo pageInfo = new PageInfo(abnormalRecordDetailList);
        return ResultGenerator.genSuccessResult(pageInfo);
    }

    /**
     * 根据异常类型、异常提交时间、提交者、解决者，返回abnormalRecordDetail
     *
     * finishStatus: 1--解决; 2--未解决
     */
    @PostMapping("/selectAbnormalRecordDetailList")
    public Result selectAbnormalRecordDetailList(@RequestParam(defaultValue = "0") Integer page,
                                                 @RequestParam(defaultValue = "0") Integer size,
                                                 String nameplate,
                                                 String orderNum,
                                                 Integer abnormalType,
                                                 String taskName,
                                                 Integer submitUser,
                                                 Integer solutionUser,
                                                 Integer finishStatus,
                                                 String queryStartTime,
                                                 String queryFinishTime) {
        PageHelper.startPage(page, size);
        List<AbnormalRecordDetail> abnormalRecordDetailList = abnormalRecordService.selectAbnormalRecordDetailList(nameplate, orderNum, abnormalType, taskName, submitUser, solutionUser, finishStatus, queryStartTime, queryFinishTime);
        PageInfo pageInfo = new PageInfo(abnormalRecordDetailList);
        return ResultGenerator.genSuccessResult(pageInfo);
    }

    /**
     * 生成 安装异常的excel表格
     */
    @PostMapping("/export")
    public Result export(
            String nameplate,
            String orderNum,
            Integer abnormalType,
            String taskName,
            Integer submitUser,
            Integer solutionUser,
            Integer finishStatus,
            String queryStartTime,
            String queryFinishTime) {
        List<AbnormalRecordDetail> list = abnormalRecordService.selectAbnormalRecordDetailList(nameplate, orderNum, abnormalType, taskName, submitUser, solutionUser, finishStatus, queryStartTime, queryFinishTime);

        HSSFWorkbook wb = null;
        FileOutputStream out = null;
        String downloadPath = "";
        /*
        返回给docker外部下载
         */
        String downloadPathForNginx = "";
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM/dd HH:mm");
        String dateString;
        try {
            //生成一个空的Excel文件
            wb = new HSSFWorkbook();
            Sheet sheet1 = wb.createSheet("sheet1");

            //设置标题行格式
            HSSFCellStyle headcellstyle = wb.createCellStyle();
            HSSFFont headfont = wb.createFont();
            headfont.setFontHeightInPoints((short) 10);
            headfont.setBold(true);
            headcellstyle.setFont(headfont);
            Row row;
            //创建行和列
            for (int r = 0; r < list.size() + 1; r++) {
                row = sheet1.createRow(r);
                for (int c = 0; c < 11; c++) {
                    row.createCell(c);
                    sheet1.getRow(0).getCell(c).setCellStyle(headcellstyle);
                }
            }
            sheet1.setColumnWidth(0, 1500);
            sheet1.setColumnWidth(1, 4000);
            sheet1.setColumnWidth(2, 6000);
            sheet1.setColumnWidth(3, 4000);
            sheet1.setColumnWidth(4, 4000);
            sheet1.setColumnWidth(5, 15000);
            sheet1.setColumnWidth(6, 4000);
            sheet1.setColumnWidth(7, 4000);
            sheet1.setColumnWidth(8, 10000);
            sheet1.setColumnWidth(9, 4000);
            sheet1.setColumnWidth(10, 4000);
            //第一行为标题
            sheet1.getRow(0).getCell(0).setCellValue("序号");
            sheet1.getRow(0).getCell(1).setCellValue("机器编号");
            sheet1.getRow(0).getCell(2).setCellValue("订单编号");
            sheet1.getRow(0).getCell(3).setCellValue("异常类型");
            sheet1.getRow(0).getCell(4).setCellValue("工序");
            sheet1.getRow(0).getCell(5).setCellValue("异常描述");
            sheet1.getRow(0).getCell(6).setCellValue("提交者");
            sheet1.getRow(0).getCell(7).setCellValue("解决者");
            sheet1.getRow(0).getCell(8).setCellValue("解决方法");
            sheet1.getRow(0).getCell(9).setCellValue("创建时间");
            sheet1.getRow(0).getCell(10).setCellValue("解决时间");

            //第二行开始，填入值
            for (int r = 0; r < list.size(); r++) {
                row = sheet1.getRow(r + 1);
                row.getCell(0).setCellValue(r + 1);
                row.getCell(1).setCellValue(list.get(r).getMachine().getNameplate());
                row.getCell(2).setCellValue(list.get(r).getOrderNum());
                row.getCell(3).setCellValue(list.get(r).getAbnormal().getAbnormalName());
                row.getCell(4).setCellValue(list.get(r).getTaskRecord().getTaskName());
                row.getCell(5).setCellValue(list.get(r).getComment());
                int userID = list.get(r).getSubmitUser();
                row.getCell(6).setCellValue(userService.findById(userID).getName());

                //安装异常时， 还不知道SolutionUser， SolutionUser是null
                if (list.get(r).getSolutionUser() != null) {
                    userID = list.get(r).getSolutionUser();
                    row.getCell(7).setCellValue(userService.findById(userID).getName());
                }
                row.getCell(8).setCellValue(list.get(r).getSolution());
                dateString = formatter.format(list.get(r).getCreateTime());
                row.getCell(9).setCellValue(dateString);
                if (list.get(r).getSolveTime() != null) {
                    dateString = formatter.format(list.get(r).getSolveTime());
                    row.getCell(10).setCellValue(dateString);
                }
            }
            downloadPath = abnoramlExcelOutputDir + "安装异常统计" + ".xls";
            downloadPathForNginx = "/excel/" + "安装异常统计" + ".xls";
            out = new FileOutputStream(downloadPath);
            wb.write(out);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        if ("".equals(downloadPath)) {
            return ResultGenerator.genFailResult("异常导出失败!");
        } else {
            return ResultGenerator.genSuccessResult(downloadPathForNginx);
        }
    }


    /**
     * 根据传入的strAbnormalRecordDetail，更新对应多表：
     * "machine_id":"", --> machine.machine_id
     * "安装是否异常":"", --> task_record.status  task状态，“1”==>未开始， “2”==>进行中，“3”==>安装完成， “4”==>质检完成，“5“===>异常
     * "异常类型":"",	--> abnormal_record.abnormal_type
     * "异常原因":"", --> abnormal_record.comment
     * "异常照片":"", -->abnormal_image.image
     * "安装完成":"",  -->   task_record.status或machine.status都可以，反正这两个表都更新
     * 注意：有外键的字段，需要上传实际存在的外键数据。
     * 一项update失败的情况下，全部update无效(事务OK)
     *
     * @param strAbnormalRecordDetail
     * @return
     */
    @PostMapping("/updateAbnormalRecordDetail")
    @Transactional
    public Result updateAbnormalRecordDetail(@RequestParam String strAbnormalRecordDetail) {
        //获取整体detail
        AbnormalRecordDetail abnormalRecordDetail1 = JSON.parseObject(strAbnormalRecordDetail, AbnormalRecordDetail.class);

        Integer abnormalRecordDetail_ID = abnormalRecordDetail1.getId();

        AbnormalRecord abnormalRecord = abnormalRecordService.findById(abnormalRecordDetail_ID);
        abnormalRecord.setAbnormalType(abnormalRecordDetail1.getAbnormalType());
        abnormalRecord.setTaskRecordId(abnormalRecordDetail1.getTaskRecordId());
        abnormalRecord.setSubmitUser(abnormalRecordDetail1.getSubmitUser());
        abnormalRecord.setSolutionUser(abnormalRecordDetail1.getSolutionUser());
        abnormalRecord.setComment(abnormalRecordDetail1.getComment());
        abnormalRecord.setSolution(abnormalRecordDetail1.getSolution());

        Abnormal abnormal = abnormalRecordDetail1.getAbnormal();
        AbnormalImage abnormalImage = abnormalRecordDetail1.getAbnormalImage();
        TaskRecord taskRecord = abnormalRecordDetail1.getTaskRecord();
        Machine machine = abnormalRecordDetail1.getMachine();

        abnormalRecordService.update(abnormalRecord);
        abnormalService.update(abnormal);
        abnormalImageService.update(abnormalImage);
        taskRecordService.update(taskRecord);
        machineService.update(machine);

        return ResultGenerator.genSuccessResult();
    }
}
