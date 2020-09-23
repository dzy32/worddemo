package com.example.worddemo.bean;

import java.util.List;

/**
 * @author ys
 * @date 2020/9/21 17:25
 */
public class TemplateData {

    /**
     * 主键
     */
    private Integer templateContentId;

    /**
     * 标题数字（1.1，1.2）
     */
    private String templateContentNumber;

    /**
     * 标题名字
     */
    private String templateContentName;

    /**
     * 这个字段没用
     */
    private String templateContentType;

    /**
     * 父级标题id
     */
    private String templateContentSuperior;

    /**
     * 模板id
     */
    private String templateId;

    /**
     * 标题等级（0为一级标题）
     */
    private Integer templateContentLevel;

    /**
     * 一级标题id
     */
    private Integer templateContentYjid;

    /**
     * 内容
     */
    List<FrameContent> frameContentList;

    public List<FrameContent> getFrameContentList() {
        return frameContentList;
    }

    public void setFrameContentList(List<FrameContent> frameContentList) {
        this.frameContentList = frameContentList;
    }

    public Integer getTemplateContentId() {
        return templateContentId;
    }

    public void setTemplateContentId(Integer templateContentId) {
        this.templateContentId = templateContentId;
    }

    public String getTemplateContentNumber() {
        return templateContentNumber;
    }

    public void setTemplateContentNumber(String templateContentNumber) {
        this.templateContentNumber = templateContentNumber;
    }

    public String getTemplateContentName() {
        return templateContentName;
    }

    public void setTemplateContentName(String templateContentName) {
        this.templateContentName = templateContentName;
    }

    public String getTemplateContentType() {
        return templateContentType;
    }

    public void setTemplateContentType(String templateContentType) {
        this.templateContentType = templateContentType;
    }

    public String getTemplateContentSuperior() {
        return templateContentSuperior;
    }

    public void setTemplateContentSuperior(String templateContentSuperior) {
        this.templateContentSuperior = templateContentSuperior;
    }

    public String getTemplateId() {
        return templateId;
    }

    public void setTemplateId(String templateId) {
        this.templateId = templateId;
    }

    public Integer getTemplateContentLevel() {
        return templateContentLevel;
    }

    public void setTemplateContentLevel(Integer templateContentLevel) {
        this.templateContentLevel = templateContentLevel;
    }

    public Integer getTemplateContentYjid() {
        return templateContentYjid;
    }

    public void setTemplateContentYjid(Integer templateContentYjid) {
        this.templateContentYjid = templateContentYjid;
    }
}
