package com.eservice.api.web;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.eservice.api.core.Result;
import com.eservice.api.core.ResultGenerator;
import com.eservice.api.model.contract.Contract;
import com.eservice.api.model.contract.ContractDetail;
import com.eservice.api.model.contract.Equipment;
import com.eservice.api.model.contract.MachineOrderWrapper;
import com.eservice.api.model.contract_sign.ContractSign;
import com.eservice.api.model.contract_sign.SignContentItem;
import com.eservice.api.model.machine.Machine;
import com.eservice.api.model.machine_order.MachineOrder;
import com.eservice.api.model.machine_order.MachineOrderDetail;
import com.eservice.api.model.order_change_record.OrderChangeRecord;
import com.eservice.api.model.order_detail.OrderDetail;
import com.eservice.api.model.order_sign.OrderSign;
import com.eservice.api.model.order_split_record.OrderSplitRecord;
import com.eservice.api.model.user.User;
import com.eservice.api.service.OrderSignService;
import com.eservice.api.service.common.CommonService;
import com.eservice.api.service.common.Constant;
import com.eservice.api.service.impl.*;
import com.eservice.api.service.mqtt.MqttMessageHelper;
import com.eservice.api.service.mqtt.ServerToClientMsg;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.interceptor.TransactionAspectSupport;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import tk.mybatis.mapper.entity.Condition;

import javax.annotation.Resource;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Class Description: xxx
 *
 * @author Wilson Hu
 * @date 2017/12/07.
 */
@RestController
@RequestMapping("/contract")
public class ContractController {
    @Resource
    private ContractServiceImpl contractService;
    @Resource
    private ContractSignServiceImpl contractSignService;
    @Resource
    private MachineOrderServiceImpl machineOrderService;
    @Resource
    private OrderDetailServiceImpl orderDetailService;
    @Resource
    private OrderSignServiceImpl orderSignService;
    @Resource
    private CommonService commonService;
    @Resource
    private OrderChangeRecordServiceImpl orderChangeRecordService;
    @Resource
    private OrderSplitRecordServiceImpl orderSplitRecordService;
    @Resource
    private MachineServiceImpl machineService;
    @Resource
    private UserServiceImpl userService;
    @Resource
    private RoleServiceImpl roleService;
    @Resource
    private MqttMessageHelper mqttMessageHelper;

    @Value("${contract_excel_output_dir}")
    private String contractOutputDir;

    @PostMapping("/add")
    @Transactional(rollbackFor = Exception.class)
    public Result add(String contract, String contractSign, String requisitionForms) {
        if (contract == null || "".equals(contract)) {
            return ResultGenerator.genFailResult("合同信息为空！");
        }
        if (contractSign == null || "".equals(contractSign)) {
            return ResultGenerator.genFailResult("合同审核初始化信息为空！");
        }
        if (requisitionForms == null || "".equals(requisitionForms)) {
            return ResultGenerator.genFailResult("订单信息为空！");
        }
        Contract contract1 = JSONObject.parseObject(contract, Contract.class);
        if (contract1 == null) {
            return ResultGenerator.genFailResult("Contract对象JSON解析失败！");
        }
        contract1.setCreateTime(new Date());
        contractService.saveAndGetID(contract1);
        Integer contractId = contract1.getId();
        ///插入contract审核记录
        ContractSign contractSignObj = new ContractSign();
        contractSignObj.setContractId(contractId);
        contractSignObj.setCreateTime(new Date());
        contractSignObj.setSignContent(contractSign);
        ///新增合同签核记录时，插入空值
        contractSignObj.setCurrentStep("");
        contractSignService.save(contractSignObj);

        //插入需求单记录
        List<MachineOrderWrapper> machineOrderWapperList = JSONObject.parseArray(requisitionForms, MachineOrderWrapper.class);
        if (machineOrderWapperList != null) {
            for (int i = 0; i < machineOrderWapperList.size(); i++) {
                OrderDetail temp = machineOrderWapperList.get(i).getOrderDetail();
                MachineOrder orderTemp = machineOrderWapperList.get(i).getMachineOrder();
                orderDetailService.saveAndGetID(temp);
                orderTemp.setOrderDetailId(temp.getId());
                orderTemp.setContractId(contractId);
                orderTemp.setStatus(Constant.ORDER_INITIAL);
                machineOrderService.saveAndGetID(orderTemp);

                //初始化需求单审核记录
                OrderSign orderSignData = machineOrderWapperList.get(i).getOrderSign();
                OrderSign orderSign = new OrderSign();
                orderSign.setSignContent(orderSignData.getSignContent());
                orderSign.setOrderId(orderTemp.getId());
                orderSign.setCreateTime(new Date());
                orderSignService.save(orderSign);
            }
        } else {
            //手动回滚事务
            TransactionAspectSupport.currentTransactionStatus().setRollbackOnly();
            return ResultGenerator.genFailResult("需求单为空！");
        }

        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/delete")
    public Result delete(@RequestParam Integer id) {
        contractService.deleteById(id);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/updateValid")
    public Result updateValid(@RequestParam Integer id) {
        if (id < 1) {
            return ResultGenerator.genFailResult("合同编号不正确，请检查！");
        }
        Contract contractObj = new Contract();
        contractObj.setId(id);
        contractObj.setIsValid((Constant.ValidEnum.INVALID.getValue()).toString());
        contractObj.setUpdateTime(new Date());
        contractService.update(contractObj);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/update")
    @Transactional(rollbackFor = Exception.class)
    public Result update(String contract, String requisitionForms) {
        Contract contract1 = JSONObject.parseObject(contract, Contract.class);
        List<MachineOrderWrapper> machineOrderWapperlist = JSONObject.parseArray(requisitionForms, MachineOrderWrapper.class);
        //先获取当前合同的所有订单
        List<MachineOrderDetail> originalOrderList = machineOrderService.selectOrder(null, contract1.getId(), null, null, null,
                null, null, null, null, null, null, false);
        ///删除该合同下，不在本次保存范围内的需求单
        for (MachineOrderDetail item : originalOrderList) {
            boolean exist = false;
            for (MachineOrderWrapper wapperItem : machineOrderWapperlist) {
                if (wapperItem.getMachineOrder().getId().equals(item.getId())) {
                    exist = true;
                    break;
                }
            }
            if (!exist) {
                //删除需求单审核记录
                OrderSign orderSign = orderSignService.findBy("orderId", item.getId());
                if (orderSign != null) {
                    orderSignService.deleteById(orderSign.getId());
                }
                //删除需求单
                machineOrderService.deleteById(item.getId());
                //删除detail
                orderDetailService.deleteById(item.getOrderDetailId());
            }
        }

        for (MachineOrderWrapper item : machineOrderWapperlist) {
            MachineOrder orderTemp = item.getMachineOrder();
            OrderChangeRecord changeRecord = item.getOrderChangeRecord();
            if (orderTemp.getId() != null && orderTemp.getId() != 0) {
                //更新，只对initial和reject状态的需求单就是更新，其他状态的需求单不做更新
                OrderDetail temp = item.getOrderDetail();
                if (orderTemp.getStatus().equals(Constant.ORDER_REJECTED)) {
                    orderTemp.setStatus(Constant.ORDER_INITIAL);
                }
                if (orderTemp.getStatus().equals(Constant.ORDER_INITIAL)) {
                    orderDetailService.update(temp);
                    machineOrderService.update(orderTemp);
                    // 在改单之后，在重新提交之前，允许修改改单原因，即：改单原因不仅仅在改单时允许修改，在上述情况下也允许修改。
                    if( null != changeRecord ) {
                        orderChangeRecordService.update(changeRecord);
                    }
                }
            } else {
                //新增
                OrderDetail temp = item.getOrderDetail();
                orderDetailService.saveAndGetID(temp);
                orderTemp.setOrderDetailId(temp.getId());
                orderTemp.setContractId(contract1.getId());
                orderTemp.setStatus(Constant.ORDER_INITIAL);
                machineOrderService.saveAndGetID(orderTemp);

                //初始化需求单审核记录
                OrderSign orderSignData = item.getOrderSign();
                OrderSign orderSign = new OrderSign();
                orderSign.setSignContent(orderSignData.getSignContent());
                orderSign.setOrderId(orderTemp.getId());
                orderSign.setCreateTime(new Date());
                orderSignService.save(orderSign);
            }
        }
        //前端只要操作了“保存”，合同的状态回到“CONTRACT_INITIAL”状态
        contract1.setStatus(Constant.CONTRACT_INITIAL);
        contract1.setUpdateTime(new Date());
        contractService.update(contract1);

        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/changeOrder")
    @Transactional(rollbackFor = Exception.class)
    public Result changeOrder(String contract, String contractSign, String requisitionForms) {
        if (contract == null || "".equals(contract)) {
            return ResultGenerator.genFailResult("合同信息为空！");
        }
        if (contractSign == null || "".equals(contractSign)) {
            return ResultGenerator.genFailResult("合同审核初始化信息为空！");
        }
        if (requisitionForms == null || "".equals(requisitionForms)) {
            return ResultGenerator.genFailResult("订单信息为空！");
        }
        Contract contract1 = JSONObject.parseObject(contract, Contract.class);
        if (contract1 == null) {
            return ResultGenerator.genFailResult("Contract对象JSON解析失败！");
        }


        //更改合同的状态为“改单”
        contract1.setStatus(Constant.CONTRACT_CHANGED);
        contract1.setUpdateTime(new Date());
        contractService.update(contract1);
        Integer contractId = contract1.getId();
        ///插入新的contract审核记录
        ContractSign contractSignObj = new ContractSign();
        contractSignObj.setContractId(contractId);
        contractSignObj.setCreateTime(new Date());
        contractSignObj.setSignContent(contractSign);
        ///插入空值
        contractSignObj.setCurrentStep("");
        contractSignService.save(contractSignObj);

        //新增的改单处理
        List<MachineOrderWrapper> machineOrderWrapperList = JSONObject.parseArray(requisitionForms, MachineOrderWrapper.class);
        for (MachineOrderWrapper orderItem : machineOrderWrapperList) {
            MachineOrder machineOrder = orderItem.getMachineOrder();
            if (machineOrder.getStatus().equals(Constant.ORDER_INITIAL) && machineOrder.getOriginalOrderId() != 0) {
                //插入新增改单项的detail
                OrderDetail temp = orderItem.getOrderDetail();
                orderDetailService.saveAndGetID(temp);
                machineOrder.setOrderDetailId(temp.getId());
                machineOrder.setContractId(contract1.getId());
                machineOrder.setStatus(Constant.ORDER_INITIAL);
                machineOrderService.saveAndGetID(machineOrder);

                //初始化需求单审核记录
                OrderSign orderSignData = orderItem.getOrderSign();
                OrderSign orderSign = new OrderSign();
                orderSign.setSignContent(orderSignData.getSignContent());
                orderSign.setOrderId(machineOrder.getId());
                orderSign.setCreateTime(new Date());
                orderSignService.save(orderSign);

                //改单记录(插入或者修改)
                OrderChangeRecord changeRecord = orderItem.getOrderChangeRecord();
                if (changeRecord.getId() == null) {
                    changeRecord.setChangeTime(new Date());
                    orderChangeRecordService.save(changeRecord);
                } else {
                    changeRecord.setChangeTime(new Date());
                    orderChangeRecordService.update(changeRecord);
                }
            }
        }

        for (MachineOrderWrapper orderItem : machineOrderWrapperList) {
            MachineOrder machineOrder = orderItem.getMachineOrder();
            //设置被改单的需求单状态(machine_order/order_sign)
            if (machineOrder.getStatus().equals(Constant.ORDER_CHANGED)) {
                //更新了被改的需求单为“改单”，持久化至数据库
                machineOrderService.update(machineOrder);
                //获取被改单对应机器，设置改单状态(machine)
                Condition tempCondition = new Condition(Machine.class);
                tempCondition.createCriteria().andCondition("order_id = ", machineOrder.getId());
                List<Machine> machineList = machineService.findByCondition(tempCondition);
                //寻找对应新需求单，比较机器数
                MachineOrder newOrder = null;
                for (MachineOrderWrapper wrapper : machineOrderWrapperList) {
                    if (wrapper.getMachineOrder().getOriginalOrderId().equals(machineOrder.getId())) {
                        newOrder = wrapper.getMachineOrder();
                        break;
                    }
                }
                if (newOrder != null) {
                    ///改单前后机器数相等或者大于原需求单数中对应的机器数;多出部分机器在审核完成以后自动添加
                    for (Machine machine : machineList) {
                        ///初始化、取消状态，直接将机器的上的需求单号直接绑定到新需求单
                        if (machine.getStatus().equals(Constant.MACHINE_INITIAL) || machine.getStatus().equals(Constant.MACHINE_CANCELED)) {
                        } else {
                            machine.setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CHANGED)));
                            ///有改单状态的机器，通知全部安装组长
                            ServerToClientMsg msg = new ServerToClientMsg();
                            msg.setOrderNum(newOrder.getOrderNum());
                            msg.setNameplate(machine.getNameplate());
                            msg.setType(ServerToClientMsg.MsgType.ORDER_CHANGE);
                            mqttMessageHelper.sendToClient(Constant.S2C_MACHINE_STATUS_CHANGE, JSON.toJSONString(msg));
                        }
                        machine.setOrderId(newOrder.getId());
                        machine.setUpdateTime(new Date());
                        machineService.update(machine);
                    }
                    /* 20180323精简了算法，对于被改的需求单，除了初始化和取消状态的机器保持状态不变，其他机器都设置为该到为状态
                    if (newOrder.getMachineNum() >= machineOrder.getMachineNum()) {
                        for (Machine machine : machineList) {
                            ///初始化状态，直接将机器的上的需求单号直接绑定到新需求单
                            if (machine.getStatus().equals(Constant.MACHINE_INITIAL)) {
                            } else {
                                machine.setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CHANGED)));
                            }
                            machine.setOrderId(newOrder.getId());
                            machine.setUpdateTime(new Date());
                            machineService.update(machine);
                        }
                    } else {
                        List<Machine> originalInitialMachine = new ArrayList<>();
                        List<Machine> originalInitialedMachine = new ArrayList<>();
                        List<Machine> originalOtherMachine = new ArrayList<>();
                        for (Machine machine : machineList) {
                            if (machine.getStatus().equals(Constant.MACHINE_CONFIGURED)
                                    ||machine.getStatus().equals(Constant.MACHINE_PLANING)
                                    || machine.getStatus().equals( Constant.MACHINE_INSTALLING)
                                    || machine.getStatus().equals(Constant.MACHINE_INSTALLED)
                                    || machine.getStatus().equals(Constant.MACHINE_SPLITED)) {
                                ///查找已配置、计划中、生产中、被拆单、生产完成的机器
                                originalInitialedMachine.add(machine);
                            } else if (machine.getStatus().equals(Constant.MACHINE_INITIAL)) {
                                ///初始化状态，未开始计划
                                originalInitialMachine.add(machine);
                            } else {
                                originalOtherMachine.add(machine);
                            }
                        }
                        int addedNum = 0;
                        //生产中的机器优先处理
                        for (int i = 0; i < originalInitialedMachine.size(); i++) {
                            if (addedNum < newOrder.getMachineNum()) {
                                originalInitialedMachine.get(i).setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CHANGED)));
                                addedNum++;
                            } else {
                                originalInitialedMachine.get(i).setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CANCELED)));
                            }
                            originalInitialedMachine.get(i).setOrderId(newOrder.getId());
                            //更新
                            machineService.update(originalInitialedMachine.get(i));
                        }
                        //未计划的机器其次处理
                        for (int i = 0; i < originalInitialMachine.size(); i++) {
                            if (addedNum < newOrder.getMachineNum()) {
                                ///TODO：是否可以保持“MACHINE_INITIAL”状态不变，因为此时机器还没有计划，也就是说没有对其设置安装流程
                                //originalInitialMachine.get(i).setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CHANGED)));
                                addedNum++;
                                originalInitialMachine.get(i).setOrderId(newOrder.getId());
                                machineService.update(originalInitialMachine.get(i));
                            } else {
                                //删除
                                machineService.deleteById(originalInitialMachine.get(i).getId());
                            }
                        }

                        //目前只有删除状态的机器（MACHINE_CANCELED）
                        for (int i = 0; i < originalOtherMachine.size(); i++) {
                            if (addedNum < newOrder.getMachineNum()) {
                                originalOtherMachine.get(i).setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CHANGED)));
                                originalOtherMachine.get(i).setOrderId(newOrder.getId());
                                machineService.update(originalInitialMachine.get(i));
                                addedNum++;
                            } else {
                                originalOtherMachine.get(i).setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CANCELED)));
                                machineService.update(originalInitialMachine.get(i));
                            }
                            ///对于删除状态的机器不做处理
//                            else {
//                                originalOtherMachine.get(i).setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_CANCELED)));
//                                machineService.update(originalInitialMachine.get(i));
//                            }
                        }
                    }
                    */
                } else {
                    ///在同一个合同中没有找到新的需求单,抛出异常
                    throw new RuntimeException();
                }
            }
        }

        return ResultGenerator.genSuccessResult();

    }

    @PostMapping("/splitOrder")
    @Transactional(rollbackFor = Exception.class)
    public Result splitOrder(String contract, String contractSign, String requisitionForms, String splitMachines) {
        if (contract == null || "".equals(contract)) {
            return ResultGenerator.genFailResult("合同信息为空！");
        }
        if (contractSign == null || "".equals(contractSign)) {
            return ResultGenerator.genFailResult("合同审核初始化信息为空！");
        }
        if (requisitionForms == null || "".equals(requisitionForms)) {
            return ResultGenerator.genFailResult("订单信息为空！");
        }
        if (splitMachines == null || "".equals(splitMachines)) {
            return ResultGenerator.genFailResult("拆单机器信息为空！");
        }

        Contract contractObj = JSONObject.parseObject(contract, Contract.class);
        if (contractObj == null || contractSign == null || requisitionForms == null || splitMachines == null) {
            return ResultGenerator.genFailResult("JSON解析失败！");
        }

        //更改合同的状态为“拆单”
        contractObj.setStatus(Constant.CONTRACT_SPLITED);
        contractObj.setUpdateTime(new Date());
        contractService.update(contractObj);
        Integer contractId = contractObj.getId();
        ///插入新的contract审核记录
        ContractSign contractSignObj = new ContractSign();
        contractSignObj.setContractId(contractId);
        contractSignObj.setCreateTime(new Date());
        contractSignObj.setSignContent(contractSign);
        ///插入空值
        contractSignObj.setCurrentStep("");
        contractSignService.save(contractSignObj);

        //新增的改单处理
        List<MachineOrderWrapper> machineOrderWrapperList = JSONObject.parseArray(requisitionForms, MachineOrderWrapper.class);
        List<Machine> splitMachineList = JSONObject.parseArray(splitMachines, Machine.class);

        for (MachineOrderWrapper orderItem : machineOrderWrapperList) {
            MachineOrder machineOrder = orderItem.getMachineOrder();
            if (machineOrder.getId() == null && machineOrder.getOriginalOrderId() != 0) {
                //插入新增改单项的detail
                OrderDetail temp = orderItem.getOrderDetail();
                orderDetailService.saveAndGetID(temp);
                machineOrder.setOrderDetailId(temp.getId());
                machineOrder.setContractId(contractObj.getId());
                machineOrder.setStatus(Constant.ORDER_INITIAL);
                machineOrderService.saveAndGetID(machineOrder);

                //初始化需求单审核记录
                OrderSign orderSignData = orderItem.getOrderSign();
                OrderSign orderSign = new OrderSign();
                orderSign.setSignContent(orderSignData.getSignContent());
                orderSign.setOrderId(machineOrder.getId());
                orderSign.setCreateTime(new Date());
                orderSignService.save(orderSign);

                //被拆分出来的机器绑定到新的需求单
                for (Machine splitMachine : splitMachineList) {
                    splitMachine.setOrderId(machineOrder.getId());
                    splitMachine.setStatus(Constant.MACHINE_SPLITED);
                    ///MQTT 有拆单状态的机器，通知全部安装组长
                    ServerToClientMsg msg = new ServerToClientMsg();
                    msg.setOrderNum(machineOrder.getOrderNum());
                    msg.setNameplate(splitMachine.getNameplate());
                    msg.setType(ServerToClientMsg.MsgType.ORDER_SPLIT);
                    mqttMessageHelper.sendToClient(Constant.S2C_MACHINE_STATUS_CHANGE, JSON.toJSONString(msg));
                    splitMachine.setUpdateTime(new Date());
                    machineService.update(splitMachine);
                }

                //拆单记录(插入或者修改)
                OrderSplitRecord splitRecord = orderItem.getOrderSplitRecord();
                if (splitRecord.getId() == null) {
                    splitRecord.setSplitTime(new Date());
                    orderSplitRecordService.save(splitRecord);
                } else {
                    splitRecord.setSplitTime(new Date());
                    orderSplitRecordService.update(splitRecord);
                }
            }
        }

        //处于拆单状态的需求单，更新状态成“ORDER_SPLIT”
        for (MachineOrderWrapper orderItem : machineOrderWrapperList) {
            MachineOrder machineOrder = orderItem.getMachineOrder();
            //TODO:同一个合同中其他为“ORDER_SPLIT”状态的需求单也会被更新，需要完善
            if (machineOrder.getStatus().equals(Constant.ORDER_SPLITED)) {
                machineOrder.setUpdateTime(new Date());
                machineOrderService.update(machineOrder);
            }
        }

        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/detail")
    public Result detail(@RequestParam Integer id) {
        Contract contract = contractService.findById(id);
        return ResultGenerator.genSuccessResult(contract);
    }

    @PostMapping("/list")
    public Result list(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size) {
        PageHelper.startPage(page, size);
        List<Contract> list = contractService.findAll();
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }

    @PostMapping("/selectContracts")
    public Result selectContracts(@RequestParam(defaultValue = "0") Integer page,
                                  @RequestParam(defaultValue = "0") Integer size,
                                  String contractNum,
                                  Integer status,
                                  String sellman,
                                  String recordUser,
                                  String roleName,
                                  String marketGroupName,
                                  String query_start_time,
                                  String query_finish_time,
                                  @RequestParam(defaultValue = "true") Boolean is_fuzzy) {
        PageHelper.startPage(page, size);
        List<ContractDetail> list = contractService.selectContracts(contractNum, status, sellman, recordUser, roleName, marketGroupName, query_start_time, query_finish_time, is_fuzzy);
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }

    @PostMapping("/startSign")
    @Transactional(rollbackFor = Exception.class)
    public Result startSign(@RequestParam Integer contractId) {
        if (contractId == null) {
            ResultGenerator.genFailResult("合同ID为空！");
        } else {
            ContractSign contractSign = contractSignService.detailByContractId(String.valueOf(contractId));
            if (contractSign == null) {
                return ResultGenerator.genFailResult("根据合同号获取合同签核信息失败！");
            } else {
                //更新合同状态为“CONTRACT_CHECKING”
                Contract contract = contractService.findById(contractId);
                if (contract == null) {
                    return ResultGenerator.genFailResult("合同编号ID无效");
                } else {
                    contract.setStatus(Constant.CONTRACT_CHECKING);
                    contract.setUpdateTime(new Date());
                    contractService.update(contract);
                }

                //设置该合同下的需求单状态，如果处于“ORDER_INITIAL”状态，则设置为“ORDER_CHECKING”
                Condition tempCondition = new Condition(MachineOrder.class);
                tempCondition.createCriteria().andCondition("contract_id = ", contractId);
                tempCondition.createCriteria().andCondition("status = ", Constant.ORDER_INITIAL);
                List<MachineOrder> orderList = machineOrderService.findByCondition(tempCondition);
                for (MachineOrder orderItem : orderList) {
                    if (orderItem.getStatus().equals(Constant.ORDER_INITIAL)) {
                        orderItem.setStatus(Constant.ORDER_CHECKING);
                        machineOrderService.update(orderItem);
                        Condition signCondition = new Condition(OrderSign.class);
                        signCondition.createCriteria().andCondition("order_id = ", orderItem.getId());
                        List<OrderSign> orderSignList = orderSignService.findByCondition(signCondition);
                        //无签核记录
                        if (orderSignList.size() == 0) {
                            throw new RuntimeException();
                        } else {
                            OrderSign sign = orderSignList.get(orderSignList.size() - 1);
                            List<SignContentItem> orderSignContentList = JSON.parseArray(sign.getSignContent(), SignContentItem.class);
                            sign.setCurrentStep(roleService.findById(orderSignContentList.get(0).getRoleId()).getRoleName());
                            orderSignService.update(sign);
                        }
                    }
                }

                //更新签核记录
//                contractSign.setUpdateTime(new Date());
//                String currentStep = commonService.getCurrentSignStep(contractId);
//                if (currentStep == null) {
//                    return ResultGenerator.genFailResult("获取当前签核steps失败！");
//                }
//                contractSign.setCurrentStep(currentStep);
//                //设置完状态后更新签核记录
//                contractSignService.update(contractSign);
            }
        }
        return ResultGenerator.genSuccessResult();
    }

    /**
     * 根据 contract_id，创建EXCEL表格，“合同评审单”+“客户需求单” 等sheet。
     * 具体内容来自 contract, contract_sign,machine_order,order_detail
     * Update: 总经理，销售，财务之外的用户，生成的excel里不显示金额信息.
     *
     * @param contractId
     * @return
     */
    @PostMapping("/buildContractExcel")
    public Result buildContractExcel(@RequestParam Integer contractId, @RequestParam String account) {
        InputStream fs = null;
        POIFSFileSystem pfs = null;
        HSSFWorkbook wb = null;
        FileOutputStream out = null;
        String downloadPath = "";
        /*
        返回给docker外部下载
         */
        String downloadPathForNginx = "";

        //只有总经理，销售，财务等用户，生成的excel里才显示金额信息. '6','7','9','14','15'
        Boolean displayPrice = false;
        User user = userService.selectByAccount(account);
        if (user != null) {
            Integer roleId = user.getRoleId();
            if ((6 == roleId)
                    || (7 == roleId)
                    || (9 == roleId)
                    || (14 == roleId)
                    || (15 == roleId)) {
                displayPrice = true;
            }
        }

        try {
            ClassPathResource resource = new ClassPathResource("empty_contract.xls");
            fs = resource.getInputStream();
            pfs = new POIFSFileSystem(fs);
            wb = new HSSFWorkbook(pfs);

            Contract contract = contractService.findById(contractId);
            if (contract == null) {
                return ResultGenerator.genFailResult("contractID not exist!");
            }

            //一个合同可能对应多个需求单
            List<Integer> machineOrderIdList = new ArrayList<Integer>();
            MachineOrder mo;
            for (int i = 0; i < machineOrderService.findAll().size(); i++) {
                mo = machineOrderService.findAll().get(i);
                if (mo.getContractId().equals(contractId)) {
                    // (已改的单，是废弃的单，不用再显示在excel里,已拆的单,因为是有效的，所以保留着)
                    if (mo.getStatus() != Constant.ORDER_CHANGED) {
                        machineOrderIdList.add(mo.getId());
                    }
                }
            }
            MachineOrder machineOrder;
            MachineOrderDetail machineOrderDetail;
            //需求单签核,一个需求单对应0个或多个签核
            List<OrderSign> orderSignList;

            //读取了模板内所有sheet1内容
            HSSFSheet sheet1 = wb.getSheetAt(0);
            //在相应的单元格进行赋值(A2)
            HSSFCell cell = sheet1.getRow(1).getCell((short) 0);
            cell.setCellValue(new HSSFRichTextString("合 同 号：" + contract.getContractNum()));
            //D2
            cell = sheet1.getRow(1).getCell((short) 3);
            SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
            SimpleDateFormat formatter2 = new SimpleDateFormat("yyyy/MM/dd");
            String dateString = formatter.format(contract.getCreateTime());
            HSSFCellStyle style = cell.getCellStyle();
            style.setWrapText(true);
            cell.setCellStyle(style);
            cell.setCellValue(new HSSFRichTextString(dateString));
            //B3
            cell = sheet1.getRow(2).getCell((short) 1);
            cell.setCellValue(new HSSFRichTextString(contract.getCustomerName()));

            //N个需求单，插入N行
            Integer machineOrderCount = machineOrderIdList.size();
            insertRow(wb, sheet1, 5, machineOrderCount);

            System.out.println("======== machineOrderCount: " + machineOrderCount);
            //订单总价 = 机器价格*台数 + 装置价格*台数 - 优惠总价
            Integer allSumOfOrder = 0;
            Integer allSumOfContract = 0;
            String machineInfo = "";
            for (int i = 0; i < machineOrderCount; i++) {
                machineOrder = machineOrderService.findById(contractId);
                machineOrderDetail = machineOrderService.getOrderAllDetail(machineOrderIdList.get(i));
                //A5,A6,A7,...品牌
                cell = sheet1.getRow(5 + i).getCell((short) 0);
                cell.setCellValue(new HSSFRichTextString(machineOrderDetail.getBrand()));

                //B5,B6,B7,...机型详细信息：机型/针数/头数/头距/x行程/y行程/剪线方式/电脑
                cell = sheet1.getRow(5 + i).getCell((short) 1);
                machineInfo = machineOrderDetail.getMachineType().getName() + "/"
                        + machineOrderDetail.getNeedleNum() + "/"
                        + machineOrderDetail.getHeadNum() + "/"
                        + machineOrderDetail.getHeadDistance() + "/"
                        + machineOrderDetail.getxDistance() + "/"
                        + machineOrderDetail.getyDistance() + "/"
                        + machineOrderDetail.getOrderDetail().getElectricTrim() + "/"
                        + machineOrderDetail.getOrderDetail().getElectricPc();
                cell.setCellValue(new HSSFRichTextString(machineInfo));

                //C5,C6,C7,...数量
                cell = sheet1.getRow(5 + i).getCell((short) 2);
                cell.setCellValue(new HSSFRichTextString(machineOrderDetail.getMachineNum().toString()));

                //D5,D6,D7,...单价
                cell = sheet1.getRow(5 + i).getCell((short) 3);
                if (displayPrice) {
                    cell.setCellValue(new HSSFRichTextString(machineOrderDetail.getMachinePrice()));
                } else {
                    cell.setCellValue(new HSSFRichTextString("/"));
                }

                //E5,E6,E7,..居间费用/台
                cell = sheet1.getRow(5 + i).getCell((short) 4);
                if (displayPrice) {
                    cell.setCellValue(new HSSFRichTextString(machineOrderDetail.getIntermediaryPrice()));
                } else {
                    cell.setCellValue(new HSSFRichTextString("/"));
                }

                //F5,F6,F7,..居间费用总计,不计入订单总价
                cell = sheet1.getRow(5 + i).getCell((short) 5);
                if (displayPrice) {
                    Integer sumOfIntermediary =Integer.parseInt(machineOrderDetail.getIntermediaryPrice()) * machineOrderDetail.getMachineNum();
                    cell.setCellValue(new HSSFRichTextString(sumOfIntermediary.toString()));
                } else {
                    cell.setCellValue(new HSSFRichTextString("/"));
                }
                //G5,G6,G7,..优惠价格/台
                cell = sheet1.getRow(5 + i).getCell((short) 6);
                if (displayPrice) {
                    cell.setCellValue(new HSSFRichTextString(machineOrderDetail.getDiscounts()));
                } else {
                    cell.setCellValue(new HSSFRichTextString("/"));
                }
                //H5,H6,H7,..优惠总计
                cell = sheet1.getRow(5 + i).getCell((short) 7);
                if (displayPrice) {
                    Integer sumOfDiscounts =Integer.parseInt(machineOrderDetail.getDiscounts()) * machineOrderDetail.getMachineNum();
                    cell.setCellValue(new HSSFRichTextString(sumOfDiscounts.toString()));
                } else {
                    cell.setCellValue(new HSSFRichTextString("/"));
                }

                //计算在合同sheet用到装置价格信息，
                JSONArray jsonArray = JSON.parseArray(machineOrderDetail.getEquipment());
                Integer equipmentCount = 0;
                //订单的设备总价
                Integer equipmentSumOfOrder = 0;
                if (null != jsonArray) {
                    equipmentCount = jsonArray.size();
                    for (int j = 0; j < equipmentCount; j++) {
                        Equipment eq = JSON.parseObject((String) jsonArray.get(j).toString(), Equipment.class);

                        //该类型的设备总价
                        int eqSum = eq.getNumber() * eq.getPrice();
                        equipmentSumOfOrder += eqSum;
                    }
                }
                // I5,I6,I7...装置总价。
                cell = sheet1.getRow(5 + i).getCell((short) 8);
                if (displayPrice) {
                    cell.setCellValue(new HSSFRichTextString(equipmentSumOfOrder.toString() ));
                } else {
                    cell.setCellValue(new HSSFRichTextString("/"));
                }
               // J5,J6,J7...订单总价
                cell = sheet1.getRow(5 + i).getCell((short) 9);
                if (displayPrice) {
                    allSumOfOrder = Integer.parseInt(machineOrderDetail.getMachinePrice()) * machineOrderDetail.getMachineNum()
                            - Integer.parseInt(machineOrderDetail.getDiscounts())*machineOrderDetail.getMachineNum()
                            + equipmentSumOfOrder;
                    allSumOfContract = allSumOfContract + allSumOfOrder;
                    cell.setCellValue(new HSSFRichTextString(allSumOfOrder.toString()));
                } else {
                    cell.setCellValue(new HSSFRichTextString("/"));
                }
            }

            Integer locationRow = 6 + machineOrderCount;
            // 合同总计
            cell = sheet1.getRow(locationRow++).getCell((short) 9);
            if (displayPrice) {
                cell.setCellValue(new HSSFRichTextString(allSumOfContract.toString()));
            } else {
                cell.setCellValue(new HSSFRichTextString("/"));
            }

            // 付款方式
            cell = sheet1.getRow(locationRow++).getCell((short) 1);
            if (displayPrice) {
                cell.setCellValue(new HSSFRichTextString(contract.getPayMethod()));
            } else {
                cell.setCellValue(new HSSFRichTextString("/"));
            }

            // 币种
            cell = sheet1.getRow(locationRow++).getCell((short) 1);
            cell.setCellValue(new HSSFRichTextString(contract.getCurrencyType()));

            // 合同交货日期
            String dateTimeString = formatter.format(contract.getContractShipDate());
            cell = sheet1.getRow(locationRow++).getCell((short) 1);
            cell.setCellValue(new HSSFRichTextString(dateTimeString));

            // 备注
            cell = sheet1.getRow(locationRow++).getCell((short) 0);
            if (displayPrice) {
                cell.setCellValue(new HSSFRichTextString(contract.getMark()));
            } else {
                cell.setCellValue(new HSSFRichTextString("/"));
            }

            // 销售员
            locationRow = locationRow + 6;
            cell = sheet1.getRow(locationRow++).getCell((short) 1);
            cell.setCellValue(new HSSFRichTextString(contract.getSellman()));

            //一个合同对应多个签核 TODO:多个签核时如何选择，contractSignService.detailByContractId 可能要改。
            ContractSign contractSign;
            // 合同审核信息，来自 contract_sign
            contractSign = contractSignService.detailByContractId(contractId.toString());
            SignContentItem signContentItem;
            List<SignContentItem> signContentItemList = JSON.parseArray(contractSign.getSignContent(), SignContentItem.class);
            OrderSign orderSign = null;

            /**
             *  根据签核内容 signContentItemList，动态填入表格。
             *  signType为 “合同签核”的，都按顺序填入表格
             */
            //合同的N个签核，插入N行
            Integer contractSignCount = signContentItemList.size();
            insertRow(wb, sheet1, locationRow, contractSignCount);
            for (int k = 0; k < contractSignCount; k++) {
                /**
                 * 合同签核的： 角色（部门）/人/时间/意见
                 */
                //1.签核角色（部门）
                int roleId = signContentItemList.get(k).getRoleId();
                //根据roleId返回角色（部门）
                String roleName = roleService.findById(roleId).getRoleName();
                cell = sheet1.getRow(locationRow).getCell((short) 0);
                cell.setCellValue(new HSSFRichTextString(roleName));
                //2.签核人
                cell = sheet1.getRow(locationRow).getCell((short) 1);
                cell.setCellValue(new HSSFRichTextString(signContentItemList.get(k).getUser()));
                //3.签核时间
                cell = sheet1.getRow(locationRow).getCell((short) 2);
                if (null != signContentItemList.get(k).getDate()) {
                    cell.setCellValue(new HSSFRichTextString(formatter2.format(signContentItemList.get(k).getDate())));
                }
                cell = sheet1.getRow(locationRow).getCell((short) 3);
                cell.setCellValue(new HSSFRichTextString("意见"));
                //4.签核意见
                cell = sheet1.getRow(locationRow++).getCell((short) 4);
                cell.setCellValue(new HSSFRichTextString(signContentItemList.get(k).getComment()));

            }
            //删除最后多余一行
            sheet1.shiftRows(locationRow + 1,
                    sheet1.getLastRowNum(),
                    -1);
            //删除需求单行的空白的多余一行
            sheet1.shiftRows(6 + machineOrderCount ,
                    sheet1.getLastRowNum(),
                    -1);

            //需求单
            //根据实际需求单数量，动态复制生成新的sheet;
            for (int i = 0; i < machineOrderCount - 1; i++) {
                // clone已经包含copy+paste
                wb.cloneSheet(1);
            }
            //调整sheet位置
            Integer sheetCount = wb.getNumberOfSheets();
            wb.setSheetOrder("Sheet3", sheetCount - 1);

            //sheet2，sheet3...,第1,2,...个需求单
            for (int i = 0; i < machineOrderCount; i++) {
                machineOrderDetail = machineOrderService.getOrderAllDetail(machineOrderIdList.get(i));
                //把sheet名称改为订单的编号
                wb.setSheetName(i + 1, machineOrderDetail.getOrderNum().replaceAll("/", "-"));

                HSSFSheet sheetX = wb.getSheetAt(1 + i);
                //在相应的单元格进行赋值
                //B2
                HSSFCell cell2 = sheetX.getRow(1).getCell((short) 1);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getSellman()));
                //D2
                cell2 = sheetX.getRow(1).getCell((short) 3);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getMaintainType()));
                //F2
                cell2 = sheetX.getRow(1).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderNum()));
                // I2
                cell2 = sheetX.getRow(1).getCell((short) 8);
                cell2.setCellValue(new HSSFRichTextString(contract.getContractNum()));

                //C3
                cell2 = sheetX.getRow(2).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getCustomer()));
                //E3
                cell2 = sheetX.getRow(2).getCell((short) 4);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getBrand()));
                //H3
                cell2 = sheetX.getRow(2).getCell((short) 7);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getMachineType().getName()));

                //C4
                cell2 = sheetX.getRow(3).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getNeedleNum()
                        + "/" + machineOrderDetail.getHeadNum().toString()));
                //E4
                cell2 = sheetX.getRow(3).getCell((short) 4);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getHeadDistance().toString()));
                //H4
                cell2 = sheetX.getRow(3).getCell((short) 7);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getxDistance()));

                //H5
                cell2 = sheetX.getRow(4).getCell((short) 7);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getyDistance()));

                //D6
                cell2 = sheetX.getRow(5).getCell((short) 3);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getSpecialTowelColor()));
                //F6
                cell2 = sheetX.getRow(5).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getSpecialTowelDaxle()));
                //H6
                cell2 = sheetX.getRow(5).getCell((short) 7);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getSpecialTowelHaxle()));
                //K6
                cell2 = sheetX.getRow(5).getCell((short) 10);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getSpecialTowelMotor()));

                //D7
                cell2 = sheetX.getRow(6).getCell((short) 3);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getSpecialTapingHead()));
                //H7
                cell2 = sheetX.getRow(6).getCell((short) 7);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getSpecialTowelNeedle()));

                //C8
                cell2 = sheetX.getRow(7).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getElectricPc()));
                //D8
                cell2 = sheetX.getRow(7).getCell((short) 3);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getElectricLanguage()));
                //F8
                cell2 = sheetX.getRow(7).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getElectricMotor()));
                //I8
                cell2 = sheetX.getRow(7).getCell((short) 8);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getElectricMotorXy()));

                //C9
                cell2 = sheetX.getRow(8).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getElectricTrim()));
                //F9
                cell2 = sheetX.getRow(8).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getElectricPower()));
                //I9 改为了 换色方式
                cell2 = sheetX.getRow(8).getCell((short) 8);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getColorChangeMode()));

                //C10
                cell2 = sheetX.getRow(9).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getElectricOil()));

                //C11
                cell2 = sheetX.getRow(10).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleSplit()));
                //F11
                cell2 = sheetX.getRow(10).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxlePanel()));
                //i11
                cell2 = sheetX.getRow(10).getCell((short) 8);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleNeedle()));

                //C12
                cell2 = sheetX.getRow(11).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleHook()));
                //f12 取消
//                cell2 = sheetX.getRow(11).getCell((short) 5);
//                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleDownCheck()));
                //i12 取消
//                cell2 = sheetX.getRow(11).getCell((short) 8);
//                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleHook()));

                //C13
                cell2 = sheetX.getRow(12).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleJump()));
                //F13
                cell2 = sheetX.getRow(12).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleUpperThread()));

                //C14
                cell2 = sheetX.getRow(13).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getAxleAddition()));

                //C15
                cell2 = sheetX.getRow(14).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getFrameworkColor()));
                //f15
                cell2 = sheetX.getRow(14).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getFrameworkPlaten()));

                //G15
                cell2 = sheetX.getRow(14).getCell((short) 6);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getFrameworkPlatenColor()));

                //i15
                cell2 = sheetX.getRow(14).getCell((short) 8);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getFrameworkRing()));

                //C16
                cell2 = sheetX.getRow(15).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getFrameworkBracket()));
                //f16
                cell2 = sheetX.getRow(15).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getFrameworkStop()));
                //i16 取消
//                cell2 = sheetX.getRow(15).getCell((short) 8);
//                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getFrameworkLight()));

                //C17
                cell2 = sheetX.getRow(16).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getDriverType()));
                //f17
                cell2 = sheetX.getRow(16).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getDriverMethod()));

                //F18
                cell2 = sheetX.getRow(17).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getDriverReelHole()));
                //F19
                cell2 = sheetX.getRow(18).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getDriverReel()));

                //C18
                cell2 = sheetX.getRow(17).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getDriverHorizonNum().toString()));
                //C19
                cell2 = sheetX.getRow(18).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getOrderDetail().getDriverVerticalNum().toString()));
                //C20
                cell2 = sheetX.getRow(19).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getPackageMethod()));

                //C21
                cell2 = sheetX.getRow(19).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getPackageMark()));

                //C20
                cell2 = sheetX.getRow(20).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getWrapMachine()));

                //C21
                cell2 = sheetX.getRow(20).getCell((short) 5);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getWrapMachineChange()));

                //C22，23 ... N 装置名称
                String str = machineOrderDetail.getEquipment();
                JSONArray jsonArray = JSON.parseArray(str);
                Integer equipmentCount = 0;
                Integer totalPriceOfOrder = 0;
                if (null != jsonArray) {
                    //该需求单的N个装置，插入N行
                    equipmentCount = jsonArray.size();
                    insertRow2(wb, sheetX, 22, equipmentCount - 1);
                    System.out.println("========order: " + machineOrderDetail.getOrderNum() + " inserted " + equipmentCount + " line");

                    for (int j = 0; j < equipmentCount; j++) {
                        Equipment eq = JSON.parseObject((String) jsonArray.get(j).toString(), Equipment.class);
                        cell2 = sheetX.getRow(22 + j).getCell((short) 0);
                        cell2.setCellValue(new HSSFRichTextString(Integer.toString(j + 1)));

                        cell2 = sheetX.getRow(22 + j).getCell((short) 1);
                        cell2.setCellValue(new HSSFRichTextString(eq.getName()));

                        cell2 = sheetX.getRow(22 + j).getCell((short) 2);
                        cell2.setCellValue(new HSSFRichTextString(eq.getNumber().toString()));

                        cell2 = sheetX.getRow(22 + j).getCell((short) 3);
                        if (displayPrice) {
                            cell2.setCellValue(new HSSFRichTextString(eq.getPrice().toString()));
                        } else {
                            cell2.setCellValue(new HSSFRichTextString("/"));
                        }
                        cell2 = sheetX.getRow(22 + j).getCell((short) 4);
                        int eqSum = eq.getNumber() * eq.getPrice();
                        totalPriceOfOrder += eqSum;
                        if (displayPrice) {
                            cell2.setCellValue(new HSSFRichTextString((Integer.toString(eqSum))));
                        } else {
                            cell2.setCellValue(new HSSFRichTextString("/"));
                        }
                    }

                } else {
                    System.out.println("========order: " + machineOrderDetail.getOrderNum() + " inserted 000 line");
                }//装置end

                //居间费用/台
                cell2 = sheetX.getRow(23 + equipmentCount).getCell((short) 1);
                if (displayPrice) {
                    cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getIntermediaryPrice()));
                }else {
                    cell2.setCellValue(new HSSFRichTextString("/"));
                }
                //居间费用总计
                cell2 = sheetX.getRow(23 + equipmentCount).getCell((short) 2);
                if (displayPrice) {
                    Integer sumOfIntermediary =Integer.parseInt(machineOrderDetail.getIntermediaryPrice()) * machineOrderDetail.getMachineNum();
                    cell2.setCellValue(new HSSFRichTextString(sumOfIntermediary.toString()));
                } else {
                    cell2.setCellValue(new HSSFRichTextString("/"));
                }
                //优惠价格/台
                cell2 = sheetX.getRow(23 + equipmentCount).getCell((short) 3);
                if (displayPrice) {
                    cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getDiscounts()));
                } else {
                    cell2.setCellValue(new HSSFRichTextString("/"));
                }
                //优惠总计
                cell2 = sheetX.getRow(23 + equipmentCount).getCell((short) 4);
                if (displayPrice) {
                    Integer sumOfDiscounts =Integer.parseInt(machineOrderDetail.getDiscounts()) * machineOrderDetail.getMachineNum();
                    cell2.setCellValue(new HSSFRichTextString(sumOfDiscounts.toString()));
                    totalPriceOfOrder -= sumOfDiscounts;
                } else {
                    cell2.setCellValue(new HSSFRichTextString("/"));
                }

                // 订机数量  这里后续用变量代替
                cell2 = sheetX.getRow(24 + equipmentCount).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getMachineNum().toString()));
                // 机器单价
                cell2 = sheetX.getRow(24 + equipmentCount).getCell((short) 3);
                if (displayPrice) {
                    cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getMachinePrice()));
                } else {
                    cell2.setCellValue(new HSSFRichTextString("/"));
                }
                // 机器总价
                Integer machineOrderSum = Integer.parseInt(machineOrderDetail.getMachinePrice()) * machineOrderDetail.getMachineNum();
                cell2 = sheetX.getRow(24 + equipmentCount).getCell((short) 4);
                if (displayPrice) {
                    cell2.setCellValue(new HSSFRichTextString(machineOrderSum.toString()));
                } else {
                    cell2.setCellValue(new HSSFRichTextString("/"));
                }

                // 需求单总价
                totalPriceOfOrder += machineOrderSum;
                cell2 = sheetX.getRow(25 + equipmentCount).getCell((short) 4);
                if (displayPrice) {
                    cell2.setCellValue(new HSSFRichTextString(totalPriceOfOrder.toString()));
                } else {
                    cell2.setCellValue(new HSSFRichTextString("/"));
                }

                // 合同的交货日期
                cell2 = sheetX.getRow(26 + equipmentCount).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(formatter2.format(contract.getContractShipDate())));

                // 计划发货日期
                cell2 = sheetX.getRow(27 + equipmentCount).getCell((short) 2);
                cell2.setCellValue(new HSSFRichTextString(formatter2.format(machineOrderDetail.getPlanShipDate())));

                // 备注
                cell2 = sheetX.getRow(28 + equipmentCount).getCell((short) 0);
                cell2.setCellValue(new HSSFRichTextString(machineOrderDetail.getMark()));

                /**
                 *  需求单审核信息，来自 order_sign， 具体有几个签核步骤，可以动态填入表格
                 */
                orderSignList = orderSignService.getOrderSignListByOrderId(machineOrderIdList.get(i));

                if (orderSignList.size() > 0) {
                    //取最后一次的签核，后续看是否需要根据时间来取最新
                    orderSign = orderSignList.get(orderSignList.size() - 1);
                    signContentItemList = JSON.parseArray(orderSign.getSignContent(), SignContentItem.class);

                    //需求单的N个签核，插入N行
                    Integer orderSignCount = signContentItemList.size();
                    insertRow2(wb, sheetX, 36 + equipmentCount, orderSignCount);
                    for (int k = 0; k < orderSignCount; k++) {
                        /**
                         * 需求单签核的： 角色（部门）/人/时间/意见
                         */
                        //1.签核角色（部门）
                        int roleId = signContentItemList.get(k).getRoleId();
                        //根据roleId返回角色（部门）
                        String roleName = roleService.findById(roleId).getRoleName();
                        cell = sheetX.getRow(36 + equipmentCount + k).getCell((short) 0);
                        cell.setCellValue(new HSSFRichTextString(roleName));
                        //2.签核人
                        cell = sheetX.getRow(36 + equipmentCount + k).getCell((short) 1);
                        cell.setCellValue(new HSSFRichTextString(signContentItemList.get(k).getUser()));
                        //3.签核时间
                        cell = sheetX.getRow(36 + equipmentCount + k).getCell((short) 2);
                        if (null != signContentItemList.get(k).getDate()) {
                            cell.setCellValue(new HSSFRichTextString(formatter2.format(signContentItemList.get(k).getDate())));
                        }
                        cell = sheetX.getRow(36 + equipmentCount + k).getCell((short) 3);
                        cell.setCellValue(new HSSFRichTextString("意见"));
                        //4.签核意见
                        cell = sheetX.getRow(36 + equipmentCount + k).getCell((short) 4);
                        cell.setCellValue(new HSSFRichTextString(signContentItemList.get(k).getComment()));
                        //合并单元格
                        sheetX.addMergedRegion(new CellRangeAddress(36 + equipmentCount + k,
                                36 + equipmentCount + k, 4, 10));
                    }
                    //最后删除多余一行
                    sheetX.shiftRows(36 + equipmentCount + orderSignCount + 1,
                            sheetX.getLastRowNum(),
                            -1);
                }
            }

            //修改模板内容导出新模板,生成路径供前端下载
            downloadPath = contractOutputDir + contract.getContractNum().replaceAll("/", "-") + ".xls";
            downloadPathForNginx = "/excel/" + contract.getContractNum().replaceAll("/", "-") + ".xls";
            out = new FileOutputStream(downloadPath);
            wb.write(out);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fs.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        if ("".equals(downloadPath)) {
            return ResultGenerator.genFailResult("生成合同文件失败!");
        } else {
            return ResultGenerator.genSuccessResult(downloadPathForNginx);
        }
    }

    private void insertRow(HSSFWorkbook wb, HSSFSheet sheet, int starRow, int rows) {
        sheet.shiftRows(starRow + 1, sheet.getLastRowNum(), rows, true, false);
        starRow = starRow - 1;

        //创建 多 行
        for (int i = 0; i < rows; i++) {
            HSSFRow sourceRow = null;
            HSSFRow targetRow = null;
            HSSFCell sourceCell = null;
            HSSFCell targetCell = null;
            short m;
            starRow = starRow + 1;
            sourceRow = sheet.getRow(starRow);
            targetRow = sheet.createRow(starRow + 1);
            targetRow.setHeight(sourceRow.getHeight());

            //创建多列
            for (m = sourceRow.getFirstCellNum(); m < 10; m++) {

                targetCell = targetRow.createCell(m);
                sourceCell = sourceRow.getCell(m);
                targetCell.setCellStyle(sourceCell.getCellStyle());
                targetCell.setCellType(sourceCell.getCellType());
            }
        }
    }

    /**
     * TODO: insertRow2和insertRow可以合并
     *
     * @param wb
     * @param sheet
     * @param starRow
     * @param rows
     */
    private void insertRow2(HSSFWorkbook wb, HSSFSheet sheet, int starRow, int rows) {
        sheet.shiftRows(starRow + 1, sheet.getLastRowNum(), rows, true, false);
        starRow = starRow - 1;

        //创建 多 行
        for (int i = 0; i < rows; i++) {
            HSSFRow sourceRow = null;
            HSSFRow targetRow = null;
            HSSFCell sourceCell = null;
            HSSFCell targetCell = null;
            short m;
            starRow = starRow + 1;
            sourceRow = sheet.getRow(starRow);
            targetRow = sheet.createRow(starRow + 1);
            targetRow.setHeight(sourceRow.getHeight());

            //创建多列
            for (m = sourceRow.getFirstCellNum(); m < 11; m++) {

                targetCell = targetRow.createCell(m);
                sourceCell = sourceRow.getCell(m);
                targetCell.setCellStyle(sourceCell.getCellStyle());
                targetCell.setCellType(sourceCell.getCellType());
            }
        }
    }

    @PostMapping("/selectAllCustomer")
    public Result selectAllCustomer(@RequestParam String name) {
        List<ContractDetail> list = contractService.selectAllCustomer(name);
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(list);
    }

    @PostMapping("/isContractExist")
    public Result isContractExist(@RequestParam String contractNum) {
        if (contractNum == null) {
            return ResultGenerator.genFailResult("请输入合同编号！");
        } else {
            Condition condition = new Condition(Contract.class);
            condition.createCriteria().andCondition("contract_num = ", contractNum);
            List<Contract> list = contractService.findByCondition(condition);
            if (list.size() == 0) {
                return ResultGenerator.genSuccessResult();
            } else {
                return ResultGenerator.genFailResult("合同编号已存在！");
            }
        }
    }
}