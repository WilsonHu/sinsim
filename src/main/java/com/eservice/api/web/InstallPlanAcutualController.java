package com.eservice.api.web;
import com.eservice.api.core.Result;
import com.eservice.api.core.ResultGenerator;
import com.eservice.api.model.install_plan_acutual.InstallPlanAcutual;
import com.eservice.api.service.InstallPlanAcutualService;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import java.util.List;

/**
* Class Description: xxx
* @author Wilson Hu
* @date 2019/07/19.
*/
@RestController
@RequestMapping("/install/plan/acutual")
public class InstallPlanAcutualController {
    @Resource
    private InstallPlanAcutualService installPlanAcutualService;

    @PostMapping("/add")
    public Result add(InstallPlanAcutual installPlanAcutual) {
        installPlanAcutualService.save(installPlanAcutual);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/delete")
    public Result delete(@RequestParam Integer id) {
        installPlanAcutualService.deleteById(id);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/update")
    public Result update(InstallPlanAcutual installPlanAcutual) {
        installPlanAcutualService.update(installPlanAcutual);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/detail")
    public Result detail(@RequestParam Integer id) {
        InstallPlanAcutual installPlanAcutual = installPlanAcutualService.findById(id);
        return ResultGenerator.genSuccessResult(installPlanAcutual);
    }

    @PostMapping("/list")
    public Result list(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size) {
        PageHelper.startPage(page, size);
        List<InstallPlanAcutual> list = installPlanAcutualService.findAll();
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }
}
