package com.eservice.api.service.common;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.eservice.api.model.contract_sign.ContractSign;
import com.eservice.api.model.contract_sign.SignContentItem;
import com.eservice.api.model.machine.Machine;
import com.eservice.api.model.machine_order.MachineOrder;
import com.eservice.api.model.order_sign.OrderSign;
import com.eservice.api.model.role.Role;
import com.eservice.api.service.impl.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import tk.mybatis.mapper.entity.Condition;

import javax.annotation.Resource;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Class Description:通用服务类
 *
 * @author Wilson Hu
 * @date 22/12/2017
 */
@Service
public class CommonService {
    @Resource
    private ContractSignServiceImpl contractSignService;
    @Resource
    private OrderSignServiceImpl orderSignService;
    @Resource
    private RoleServiceImpl roleService;
    @Resource
    private MachineOrderServiceImpl machineOrderService;
    @Resource
    private MachineServiceImpl machineService;

    /**
     * 用于返回对应合同的所有签核记录，每一次提交审核以后，都需要通过该API获取所有审核内容，再设置审核状态
     * @param contractId
     * @return result, 如果签核流程不存在返回null， 未结束则返回正在签核中的名称，结束则返回“FINISHED”
     */
    public String getCurrentSignStep(Integer contractId) {
        String result = null;

        //TODO:通过合同号找到有效的订单签核记录
        List<OrderSign> orderSignList = orderSignService.getValidOrderSigns(contractId);
        //找到有效的合同签核记录
        ContractSign contractSign1 = contractSignService.detailByContractId(String.valueOf(contractId));
        HashMap<Integer, List<SignContentItem>> signContentSortByNumberMap = new HashMap<>();
        for (OrderSign os: orderSignList) {
            String contentStr = os.getSignContent();
            List<SignContentItem> orderSignContentList = JSON.parseArray(contentStr, SignContentItem.class);
            for (SignContentItem item: orderSignContentList) {
                if(signContentSortByNumberMap.get(item.getNumber()) != null) {
                    signContentSortByNumberMap.get(item.getNumber()).add(item);
                }else {
                    List<SignContentItem> list = new ArrayList<>();
                    list.add(item);
                    signContentSortByNumberMap.put(item.getNumber(), list);
                }
            }
        }
        List<SignContentItem> contractSignContentList = JSONObject.parseArray(contractSign1.getSignContent(), SignContentItem.class);
        for (SignContentItem item: contractSignContentList) {
            if(signContentSortByNumberMap.get(item.getNumber()) != null) {
                signContentSortByNumberMap.get(item.getNumber()).add(item);
            }else {
                List<SignContentItem> list = new ArrayList<>();
                list.add(item);
                signContentSortByNumberMap.put(item.getNumber(), list);
            }
        }

        //签核流程不存在
        if(signContentSortByNumberMap.size() == 0) {
            return result;
        }
        // 将map.entrySet()转换成list
        List<Map.Entry<Integer, List<SignContentItem>>> list = new ArrayList<>(signContentSortByNumberMap.entrySet());
        // 通过比较器来实现排序
        Collections.sort(list, new Comparator<Map.Entry<Integer, List<SignContentItem>>>() {
            @Override
            public int compare(Map.Entry<Integer, List<SignContentItem>> o1, Map.Entry<Integer, List<SignContentItem>> o2) {
                // 升序排序
                return o1.getKey().compareTo(o2.getKey());
            }
        });

        Iterator<Map.Entry<Integer, List<SignContentItem>>>  entries = signContentSortByNumberMap.entrySet().iterator();
        boolean currentStepFound = false;
        while (entries.hasNext() && !currentStepFound) {
            Map.Entry<Integer, List<SignContentItem>> entry = entries.next();
            Integer key = entry.getKey();
            List<SignContentItem> value = entry.getValue();
            for (SignContentItem item: value) {
                if(item.getResult().equals(Constant.SIGN_REJECT) || item.getComment() == null || "".equals(item.getComment())
                        || item.getUser() == null || "".equals(item.getUser())) {
                    List<Role> roleList = roleService.findAll();
                    for (Role role: roleList) {
                        if(role.getId().equals(item.getRoleId())) {
                            result = role.getRoleName();
                        }
                    }
                    currentStepFound = true;
                }
            }
        }
        //如果没有找到，表示该签核流程已经完成，此时返回一个特殊的String给调用者
        if(!currentStepFound) {
            result = Constant.SIGN_FINISHED;
        }

        return result;
    }

    /**
     * 根据合同编号对应的需求单的机器
     * @param contractId
     */
    public void createMachineByContractId(Integer contractId) {
        Condition condition = new Condition(MachineOrder.class);
        condition.createCriteria().andCondition("contract_id = ", contractId);
        List<MachineOrder> orderList = machineOrderService.findByCondition(condition);
        for (MachineOrder orderItem: orderList) {
            //选取有效需求单，无效需求单对应的机器数不cover在内
            if(orderItem.getStatus() < Constant.ORDER_CHANGED) {
                Condition tempCondition = new Condition(Machine.class);
                tempCondition.createCriteria().andCondition("order_id = ", orderItem.getId());
                List<Machine> machineExistList = machineService.findByCondition(tempCondition);
                int haveToCreate = orderItem.getMachineNum() - machineExistList.size();
                int i=1;
                while (i <= haveToCreate) {
                    Machine machine = new Machine();
                    machine.setMachineId(Utils.createMachineBasicId() + i);
                    machine.setOrderId(orderItem.getId());
                    machine.setMachineType(orderItem.getMachineType());
                    machine.setStatus(Byte.parseByte(String.valueOf(Constant.MACHINE_INITIAL)));
                    machine.setCreateTime(new Date());
                    machineService.save(machine);
                    i++;
                }
            }
        }
    }

    /**
     * @param path 保存文件的总路径
     * @param file 文件名称
     * @param machineID 机器的具体ID
     * @param type 文件类型 （不同类型文件名生成规则不同）
     * 最终生成类似：machineID123_Abnormal_2018-01-10-11-15-56.png
     * @return 文件路径
     */
    public String saveFile(String path, MultipartFile file, String machineID, int type) {
        String targetFileName = null;
        if(path != null) {
            if (!file.isEmpty()) {
                try {
                    // 这里只是简单例子，文件直接输出到path路径下。
                    // 实际项目中，文件需要输出到指定位置，需要在增加代码处理。
                    // 还有关于文件格式限制、文件大小限制，详见：中配置。

                    //取后缀名
                    String fileName = file.getOriginalFilename();
                    String suffixName = fileName.substring(fileName.lastIndexOf("."));

                    Date date = new Date();
                    SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
                    String dateStr = formatter.format(date);

                    String fileType ;
                    if(Constant.ABNORMAL_IMAGE == type){
                        fileType = "Abnormal";
                    } else if(Constant.QUALITY_IMAGE == type) {
                        fileType = "Quality";
                    } else if(Constant.LOADING_FILE == type) {
                        fileType = "LoadingFile";
                    } else {
                        fileType = "";//return targetFileName;//"UnknownFileTypeError";
                    }

                    targetFileName = path + machineID + "_" + fileType + "_" + dateStr + suffixName;
                    BufferedOutputStream out = new BufferedOutputStream(
                            new FileOutputStream(new File(targetFileName)));
                    out.write(file.getBytes());
                    out.flush();
                    out.close();

                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                    return null;
                } catch (IOException e) {
                    e.printStackTrace();
                    return null;
                }
            }
        }
        return targetFileName;
    }
}
