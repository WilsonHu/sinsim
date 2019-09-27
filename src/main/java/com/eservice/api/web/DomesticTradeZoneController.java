package com.eservice.api.web;
import com.eservice.api.core.Result;
import com.eservice.api.core.ResultGenerator;
import com.eservice.api.model.domestic_trade_zone.DomesticTradeZone;
import com.eservice.api.service.DomesticTradeZoneService;
import com.eservice.api.service.impl.DomesticTradeZoneServiceImpl;
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
* @date 2019/07/09.
*/
@RestController
@RequestMapping("/domestic/trade/zone")
public class DomesticTradeZoneController {
    @Resource
    private DomesticTradeZoneServiceImpl domesticTradeZoneService;

    @PostMapping("/add")
    public Result add(DomesticTradeZone domesticTradeZone) {
        domesticTradeZoneService.save(domesticTradeZone);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/delete")
    public Result delete(@RequestParam Integer id) {
        domesticTradeZoneService.deleteById(id);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/update")
    public Result update(DomesticTradeZone domesticTradeZone) {
        domesticTradeZoneService.update(domesticTradeZone);
        return ResultGenerator.genSuccessResult();
    }

    @PostMapping("/detail")
    public Result detail(@RequestParam Integer id) {
        DomesticTradeZone domesticTradeZone = domesticTradeZoneService.findById(id);
        return ResultGenerator.genSuccessResult(domesticTradeZone);
    }

    @PostMapping("/list")
    public Result list(@RequestParam(defaultValue = "0") Integer page, @RequestParam(defaultValue = "0") Integer size) {
        PageHelper.startPage(page, size);
        List<DomesticTradeZone> list = domesticTradeZoneService.findAll();
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(list);
    }

    // 根据账号返回 该账号负责的全部内贸区域
    @PostMapping("/getDomesticTradeZone")
    public Result getDomesticTradeZone(@RequestParam String account) {
        List<DomesticTradeZone> list = domesticTradeZoneService.getDomesticTradeZone(account);
        PageInfo pageInfo = new PageInfo(list);
        return ResultGenerator.genSuccessResult(pageInfo);
    }
}
