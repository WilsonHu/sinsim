package com.eservice.api.model.contract_sign;

import java.util.Date;
import javax.persistence.*;

@Table(name = "contract_sign")
public class ContractSign {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Integer id;

    /**
     * 合同ID
     */
    @Column(name = "contract_id")
    private Integer contractId;

    /**
     * 当前进行中的签核环节（来至于role_name）
     */
    @Column(name = "current_step")
    private String currentStep;

    /**
     * 创建时间
     */
    @Column(name = "create_time")
    private Date createTime;

    /**
     * 更新时间
     */
    @Column(name = "update_time")
    private Date updateTime;

    /**
     * 签核内容，以json格式的数组形式存放, 所有项完成后更新status为完成.
     * [{"number":1,"roleId":7,"signType":"合同签核","date":"","user":"","comment":""},
     * {"number":4,"roleId":13,"signType":"合同签核","date":"","user":"","comment":""},
     * {"number":5,"roleId":14,"signType":"合同签核","date":"","user":"","comment":""},
     * {"number":6,"roleId":6,"signType":"合同签核","date":"","user":"","comment":""}]
     * */
    @Column(name = "sign_content")
    private String signContent;

    /**
     * @return id
     */
    public Integer getId() {
        return id;
    }

    /**
     * @param id
     */
    public void setId(Integer id) {
        this.id = id;
    }

    /**
     * 获取合同ID
     *
     * @return contract_id - 合同ID
     */
    public Integer getContractId() {
        return contractId;
    }

    /**
     * 设置合同ID
     *
     * @param contractId 合同ID
     */
    public void setContractId(Integer contractId) {
        this.contractId = contractId;
    }

    /**
     * 获取当前进行中的签核环节（来至于role_name）
     *
     * @return current_step - 当前进行中的签核环节（来至于role_name）
     */
    public String getCurrentStep() {
        return currentStep;
    }

    /**
     * 设置当前进行中的签核环节（来至于role_name）
     *
     * @param currentStep 当前进行中的签核环节（来至于role_name）
     */
    public void setCurrentStep(String currentStep) {
        this.currentStep = currentStep;
    }

    /**
     * 获取创建时间
     *
     * @return create_time - 创建时间
     */
    public Date getCreateTime() {
        return createTime;
    }

    /**
     * 设置创建时间
     *
     * @param createTime 创建时间
     */
    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    /**
     * 获取更新时间
     *
     * @return update_time - 更新时间
     */
    public Date getUpdateTime() {
        return updateTime;
    }

    /**
     * 设置更新时间
     *
     * @param updateTime 更新时间
     */
    public void setUpdateTime(Date updateTime) {
        this.updateTime = updateTime;
    }

    /**
     * 获取签核内容，以json格式的数组形式存放, 所有项完成后更新status为完成.[ 
    {"step_number":1, "role_id": 1, "role_name":"销售经理"，“person”：“张三”，”comment“: "同意"，"resolved":1,”update_time“:"2017-11-05 12:08:55"},
    {"step_number":1,"role_id":2, "role_name":"财务部"，“person”：“李四”，”comment“: "同意，但是部分配件需要新设计"，"resolved":0, ”update_time“:"2017-11-06 12:08:55"}
]
     *
     * @return sign_content - 签核内容，以json格式的数组形式存放, 所有项完成后更新status为完成.[ 
    {"step_number":1, "role_id": 1, "role_name":"销售经理"，“person”：“张三”，”comment“: "同意"，"resolved":1,”update_time“:"2017-11-05 12:08:55"},
    {"step_number":1,"role_id":2, "role_name":"财务部"，“person”：“李四”，”comment“: "同意，但是部分配件需要新设计"，"resolved":0, ”update_time“:"2017-11-06 12:08:55"}
]
     */
    public String getSignContent() {
        return signContent;
    }

    /**
     * 设置签核内容，以json格式的数组形式存放, 所有项完成后更新status为完成.[ 
    {"step_number":1, "role_id": 1, "role_name":"销售经理"，“person”：“张三”，”comment“: "同意"，"resolved":1,”update_time“:"2017-11-05 12:08:55"},
    {"step_number":1,"role_id":2, "role_name":"财务部"，“person”：“李四”，”comment“: "同意，但是部分配件需要新设计"，"resolved":0, ”update_time“:"2017-11-06 12:08:55"}
]
     *
     * @param signContent 签核内容，以json格式的数组形式存放, 所有项完成后更新status为完成.[ 
    {"step_number":1, "role_id": 1, "role_name":"销售经理"，“person”：“张三”，”comment“: "同意"，"resolved":1,”update_time“:"2017-11-05 12:08:55"},
    {"step_number":1,"role_id":2, "role_name":"财务部"，“person”：“李四”，”comment“: "同意，但是部分配件需要新设计"，"resolved":0, ”update_time“:"2017-11-06 12:08:55"}
]
     */
    public void setSignContent(String signContent) {
        this.signContent = signContent;
    }
}