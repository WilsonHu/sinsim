package com.eservice.api.model.machine;

import java.util.Date;
import javax.persistence.*;

public class MachineInfo extends Machine {

    /**
     * 订单编号
     */
    private String orderNum;


    /**
     * 合同号对应ID
     */
    private Integer contractId;


    /**
     * 合同编号
     */
    private String contractNum;

    /*
    流程进度ID
    标识是否已配置过流程
     */
    private String processRecordId;

    @Column(name = "process_id")
    private Integer processId;

    /**
     * 安装流程的link数据,格式参考linkDataArray
     */
    @Column(name = "link_data")
    private String linkData;

    /**
     * 安装流程的node数据，格式参考nodeDataArray
     */
    @Column(name = "node_data")
    private String nodeData;

    /*
    gets sets orderNum
     */
    public String getOrderNum() {
        return orderNum;
    }

    public void setOrderNum(String orderNum) {
        this.orderNum = orderNum;
    }


    public Integer getContractId() {
        return contractId;
    }

    public void setContractId(Integer contractId) {
        this.contractId = contractId;
    }

    public String getContractNum() {
        return contractNum;
    }

    public void setContractNum(String contractNum) {
        this.contractNum = contractNum;
    }

    //processRecordId
    public String getProcessRecordId() {
        return processRecordId;
    }

    public void setProcessRecordId(String processRecordId) {
        this.processRecordId = processRecordId;
    }

    /**
     * 获取对应的模板（process）的ID
     *
     * @return process_id - 对应的模板（process）的ID
     */
    public Integer getProcessId() {
        return processId;
    }

    /**
     * 设置对应的模板（process）的ID
     *
     * @param processId 对应的模板（process）的ID
     */
    public void setProcessId(Integer processId) {
        this.processId = processId;
    }

    /**
     * 获取安装流程的link数据,格式参考linkDataArray
     *
     * @return link_data - 安装流程的link数据,格式参考linkDataArray
     */
    public String getLinkData() {
        return linkData;
    }

    /**
     * 设置安装流程的link数据,格式参考linkDataArray
     *
     * @param linkData 安装流程的link数据,格式参考linkDataArray
     */
    public void setLinkData(String linkData) {
        this.linkData = linkData;
    }

    /**
     * 获取安装流程的node数据，格式参考nodeDataArray
     *
     * @return node_data - 安装流程的node数据，格式参考nodeDataArray
     */
    public String getNodeData() {
        return nodeData;
    }

    /**
     * 设置安装流程的node数据，格式参考nodeDataArray
     *
     * @param nodeData 安装流程的node数据，格式参考nodeDataArray
     */
    public void setNodeData(String nodeData) {
        this.nodeData = nodeData;
    }
}
