package cc.whohow.excel;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonPropertyDescription;

import java.util.Date;

public class DataModel {
    private Long id;
    @JsonProperty(index = 1)
    @JsonPropertyDescription("幼儿园")
    private String name;
    @JsonPropertyDescription("省")
    private String province;
    @JsonPropertyDescription("市")
    private String city;
    @JsonPropertyDescription("县")
    private String county;
    @JsonPropertyDescription("园长/法人")
    private String leader;
    @JsonPropertyDescription("电话")
    @JsonProperty(index = 6)
    private String phone;
    @JsonPropertyDescription("运营人员")
    private String contactName;
    @JsonPropertyDescription("电话")
    @JsonProperty(index = 8)
    private String contactPhone;
    @JsonPropertyDescription("地址")
    private String address;
    @JsonPropertyDescription("授权")
    private String weixinAppId;
    @JsonPropertyDescription("激活")
    private String activation;
    @JsonPropertyDescription("总积分")
    private String score;
    @JsonPropertyDescription("活跃度")
    private String activeRate;
    @JsonPropertyDescription("创建时间")
    private Date createTime;
    @JsonPropertyDescription("审核时间")
    private Date auditTime;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getProvince() {
        return province;
    }

    public void setProvince(String province) {
        this.province = province;
    }

    public String getCity() {
        return city;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public String getCounty() {
        return county;
    }

    public void setCounty(String county) {
        this.county = county;
    }

    public String getLeader() {
        return leader;
    }

    public void setLeader(String leader) {
        this.leader = leader;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getContactName() {
        return contactName;
    }

    public void setContactName(String contactName) {
        this.contactName = contactName;
    }

    public String getContactPhone() {
        return contactPhone;
    }

    public void setContactPhone(String contactPhone) {
        this.contactPhone = contactPhone;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getWeixinAppId() {
        return weixinAppId;
    }

    public void setWeixinAppId(String weixinAppId) {
        this.weixinAppId = weixinAppId;
    }

    public String getActivation() {
        return activation;
    }

    public void setActivation(String activation) {
        this.activation = activation;
    }

    public String getScore() {
        return score;
    }

    public void setScore(String score) {
        this.score = score;
    }

    public String getActiveRate() {
        return activeRate;
    }

    public void setActiveRate(String activeRate) {
        this.activeRate = activeRate;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    public Date getAuditTime() {
        return auditTime;
    }

    public void setAuditTime(Date auditTime) {
        this.auditTime = auditTime;
    }

    @Override
    public String toString() {
        return "DataModel{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", province='" + province + '\'' +
                ", city='" + city + '\'' +
                ", county='" + county + '\'' +
                ", leader='" + leader + '\'' +
                ", phone='" + phone + '\'' +
                ", contactName='" + contactName + '\'' +
                ", contactPhone='" + contactPhone + '\'' +
                ", address='" + address + '\'' +
                ", weixinAppId='" + weixinAppId + '\'' +
                ", activation='" + activation + '\'' +
                ", score='" + score + '\'' +
                ", activeRate='" + activeRate + '\'' +
                ", createTime=" + createTime +
                ", auditTime=" + auditTime +
                '}';
    }
}
