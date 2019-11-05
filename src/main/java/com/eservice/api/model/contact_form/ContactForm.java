package com.eservice.api.model.contact_form;

import java.util.Date;
import javax.persistence.*;

@Table(name = "contact_form")
public class ContactForm {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Integer id;

    /**
     * 联系单类型： 变更联系单、工作联系单、配件申请联系单
     */
    @Column(name = "contact_type")
    private String contactType;

    /**
     * 联系单的编号,可选
     */
    private String num;

    /**
     * 对应的订单ID,可选
     */
    @Column(name = "order_id")
    private Integer orderId;

    /**
     * 提出申请的部门
     */
    @Column(name = "applicant_department")
    private String applicantDepartment;

    /**
     * 申请人
     */
    @Column(name = "applicant_person")
    private String applicantPerson;

    /**
     * 申请日期
     */
    @Column(name = "create_date")
    private Date createDate;

    /**
     * 希望完成的日期
     */
    @Column(name = "hope_date")
    private Date hopeDate;

    /**
     * 联络主题、变更理由/主题
     */
    @Column(name = "contact_title")
    private String contactTitle;

    /**
     * 联络内容
     */
    @Column(name = "contact_content")
    private String contactContent;

    private String status;

    /**
     * JSON格式，定义要参与审批的部门和顺序。
     */
    @Column(name = "defined_sign_process")
    private String definedSignProcess;

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
     * 获取联系单类型： 变更联系单、工作联系单、配件申请联系单
     *
     * @return contact_type - 联系单类型： 变更联系单、工作联系单、配件申请联系单
     */
    public String getContactType() {
        return contactType;
    }

    /**
     * 设置联系单类型： 变更联系单、工作联系单、配件申请联系单
     *
     * @param contactType 联系单类型： 变更联系单、工作联系单、配件申请联系单
     */
    public void setContactType(String contactType) {
        this.contactType = contactType;
    }

    /**
     * 获取联系单的编号,可选
     *
     * @return num - 联系单的编号,可选
     */
    public String getNum() {
        return num;
    }

    /**
     * 设置联系单的编号,可选
     *
     * @param num 联系单的编号,可选
     */
    public void setNum(String num) {
        this.num = num;
    }

    /**
     * 获取对应的订单ID,可选
     *
     * @return order_id - 对应的订单ID,可选
     */
    public Integer getOrderId() {
        return orderId;
    }

    /**
     * 设置对应的订单ID,可选
     *
     * @param orderId 对应的订单ID,可选
     */
    public void setOrderId(Integer orderId) {
        this.orderId = orderId;
    }

    /**
     * 获取提出申请的部门
     *
     * @return applicant_department - 提出申请的部门
     */
    public String getApplicantDepartment() {
        return applicantDepartment;
    }

    /**
     * 设置提出申请的部门
     *
     * @param applicantDepartment 提出申请的部门
     */
    public void setApplicantDepartment(String applicantDepartment) {
        this.applicantDepartment = applicantDepartment;
    }

    /**
     * 获取申请人
     *
     * @return applicant_person - 申请人
     */
    public String getApplicantPerson() {
        return applicantPerson;
    }

    /**
     * 设置申请人
     *
     * @param applicantPerson 申请人
     */
    public void setApplicantPerson(String applicantPerson) {
        this.applicantPerson = applicantPerson;
    }

    /**
     * 获取申请日期
     *
     * @return create_date - 申请日期
     */
    public Date getCreateDate() {
        return createDate;
    }

    /**
     * 设置申请日期
     *
     * @param createDate 申请日期
     */
    public void setCreateDate(Date createDate) {
        this.createDate = createDate;
    }

    /**
     * 获取希望完成的日期
     *
     * @return hope_date - 希望完成的日期
     */
    public Date getHopeDate() {
        return hopeDate;
    }

    /**
     * 设置希望完成的日期
     *
     * @param hopeDate 希望完成的日期
     */
    public void setHopeDate(Date hopeDate) {
        this.hopeDate = hopeDate;
    }

    /**
     * 获取联络主题、变更理由/主题
     *
     * @return contact_title - 联络主题、变更理由/主题
     */
    public String getContactTitle() {
        return contactTitle;
    }

    /**
     * 设置联络主题、变更理由/主题
     *
     * @param contactTitle 联络主题、变更理由/主题
     */
    public void setContactTitle(String contactTitle) {
        this.contactTitle = contactTitle;
    }

    /**
     * 获取联络内容
     *
     * @return contact_content - 联络内容
     */
    public String getContactContent() {
        return contactContent;
    }

    /**
     * 设置联络内容
     *
     * @param contactContent 联络内容
     */
    public void setContactContent(String contactContent) {
        this.contactContent = contactContent;
    }

    /**
     * @return status
     */
    public String getStatus() {
        return status;
    }

    /**
     * @param status
     */
    public void setStatus(String status) {
        this.status = status;
    }

    /**
     * 获取JSON格式，定义要参与审批的部门和顺序。
     *
     * @return defined_sign_process - JSON格式，定义要参与审批的部门和顺序。
     */
    public String getDefinedSignProcess() {
        return definedSignProcess;
    }

    /**
     * 设置JSON格式，定义要参与审批的部门和顺序。
     *
     * @param definedSignProcess JSON格式，定义要参与审批的部门和顺序。
     */
    public void setDefinedSignProcess(String definedSignProcess) {
        this.definedSignProcess = definedSignProcess;
    }
}