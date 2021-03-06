package com.eservice.api.web;
import com.alibaba.fastjson.JSONObject;
import com.eservice.api.core.Result;
import com.eservice.api.core.ResultGenerator;
import com.eservice.api.model.role.Role;
import com.eservice.api.service.RoleService;
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
* @date 2017/11/14.
*/
@RestController
@RequestMapping("/role")
public class RoleController {
    @Resource
    private RoleService roleService;

    @PostMapping("/add")
    public Result add(String role) {
        Role role1 = JSONObject.parseObject(role, Role.class);
        roleService.save(role1);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/delete")
    public Result delete(@RequestParam Integer id) {
        //TODO:检查签核流程中（sign_process）是否有该角色，否则无法删除
        roleService.deleteById(id);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/update")
    public Result update(String role) {
        //TODO:检查签核流程中（sign_process）是否有该角色,否则无法修改
        Role role1 = JSONObject.parseObject(role, Role.class);
        roleService.update(role1);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/detail")
    public Result detail(@RequestParam Integer id) {
        Role role = roleService.findById(id);
        return ResultGenerator.genSuccessResult(role);
    }

    @PostMapping("/list")
    public Result list(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size) {
        PageHelper.startPage(page, size);
        List<Role> list = roleService.findAll();
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }
}
