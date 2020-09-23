package com.example.worddemo.bean;

/**
 * @author ys
 * @date 2020/9/21 17:50
 */
public class FrameContent {

    /**
     * 主键
     */
    private Integer templateFrameId;

    /**
     * 框的类型 （1为单选框，2为文本框，3为上传文件狂）
     */
    private String templateFrameType;

    /**
     * 内容对应id
     */
    private Integer templateContentId;

    /**
     * 模板id
     */
    private Integer templateId;

    /**
     *内容
     */
    private String content;

    public Integer getTemplateFrameId() {
        return templateFrameId;
    }

    public void setTemplateFrameId(Integer templateFrameId) {
        this.templateFrameId = templateFrameId;
    }

    public String getTemplateFrameType() {
        return templateFrameType;
    }

    public void setTemplateFrameType(String templateFrameType) {
        this.templateFrameType = templateFrameType;
    }

    public Integer getTemplateContentId() {
        return templateContentId;
    }

    public void setTemplateContentId(Integer templateContentId) {
        this.templateContentId = templateContentId;
    }

    public Integer getTemplateId() {
        return templateId;
    }

    public void setTemplateId(Integer templateId) {
        this.templateId = templateId;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }
}
