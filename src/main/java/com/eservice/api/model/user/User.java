package com.eservice.api.model.user;

import javax.persistence.*;

public class User {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Integer id;

    /**
     * 账号
     */
    private String account;

    /**
     * 用户姓名
     */
    private String name;

    /**
     * 角色ID
     */
    @Column(name = "role_id")
    private Integer roleId;

    /**
     * 密码
     */
    private String password;

    /**
     * 所在安装组group ID，可以为空(其他部门人员)
     */
    @Column(name = "group_id")
    private Integer groupId;

    /**
     * 销售部信息（区分外贸、内贸一部、内贸二部）
     */
    @Column(name = "market_group_name")
    private String marketGroupName;

    /**
     * 员工是否在职，“1”==>在职, “0”==>离职
     */
    private Byte valid;

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
     * 获取账号
     *
     * @return account - 账号
     */
    public String getAccount() {
        return account;
    }

    /**
     * 设置账号
     *
     * @param account 账号
     */
    public void setAccount(String account) {
        this.account = account;
    }

    /**
     * 获取用户姓名
     *
     * @return name - 用户姓名
     */
    public String getName() {
        return name;
    }

    /**
     * 设置用户姓名
     *
     * @param name 用户姓名
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * 获取角色ID
     *
     * @return role_id - 角色ID
     */
    public Integer getRoleId() {
        return roleId;
    }

    /**
     * 设置角色ID
     *
     * @param roleId 角色ID
     */
    public void setRoleId(Integer roleId) {
        this.roleId = roleId;
    }

    /**
     * 获取密码
     *
     * @return password - 密码
     */
    public String getPassword() {
        return password;
    }

    /**
     * 设置密码
     *
     * @param password 密码
     */
    public void setPassword(String password) {
        this.password = password;
    }

    /**
     * 获取所在安装组group ID，可以为空(其他部门人员)
     *
     * @return group_id - 所在安装组group ID，可以为空(其他部门人员)
     */
    public Integer getGroupId() {
        return groupId;
    }

    /**
     * 设置所在安装组group ID，可以为空(其他部门人员)
     *
     * @param groupId 所在安装组group ID，可以为空(其他部门人员)
     */
    public void setGroupId(Integer groupId) {
        this.groupId = groupId;
    }

    public String getMarketGroupName() {
        return marketGroupName;
    }

    public void setMarketGroupName(String marketGroupName) {
        this.marketGroupName = marketGroupName;
    }

    /**
     * 获取员工是否在职，“1”==>在职, “0”==>离职
     *
     * @return valid - 员工是否在职，“1”==>在职, “0”==>离职
     */
    public Byte getValid() {
        return valid;
    }

    /**
     * 设置员工是否在职，“1”==>在职, “0”==>离职
     *
     * @param valid 员工是否在职，“1”==>在职, “0”==>离职
     */
    public void setValid(Byte valid) {
        this.valid = valid;
    }

    /**
     * 增加外网登录管理，管理员可自由配置账号的登录外网权限。
     * 外网时，检查用户权限是否允许外网登录; 0:不允许，1:允许。
     */
    private Byte extranetPermit;

    public Byte getExtranetPermit() {
        return extranetPermit;
    }

    public void setExtranetPermit(Byte extranetPermit) {
        this.extranetPermit = extranetPermit;
    }

    //是否接收推送消息,默认1开启，如果有人不想收，可以设为0即关闭。
    private Byte acceptWxMsg;

    public Byte getAcceptWxMsg() {
        return acceptWxMsg;
    }

    public void setAcceptWxMsg(Byte acceptWxMsg) {
        this.acceptWxMsg = acceptWxMsg;
    }
}