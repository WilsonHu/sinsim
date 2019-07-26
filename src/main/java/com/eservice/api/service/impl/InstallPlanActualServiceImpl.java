package com.eservice.api.service.impl;

import com.eservice.api.dao.InstallPlanActualMapper;
import com.eservice.api.model.install_plan_actual.InstallPlanActual;
import com.eservice.api.model.install_plan_actual.InstallPlanActualDetails;
import com.eservice.api.service.InstallPlanActualService;
import com.eservice.api.core.AbstractService;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.List;


/**
* Class Description: xxx
* @author Wilson Hu
* @date 2019/07/19.
*/
@Service
@Transactional
public class InstallPlanActualServiceImpl extends AbstractService<InstallPlanActual> implements InstallPlanActualService {
    @Resource
    private InstallPlanActualMapper installPlanActualMapper;

    public List<InstallPlanActualDetails> selectInstallPlanActualDetails(String orderNum,
                                                                         String nameplate,
                                                                         String installGroupName,
                                                                         String type) {
        return installPlanActualMapper.selectInstallPlanActualDetails(orderNum, nameplate, installGroupName, type);
    }
}
