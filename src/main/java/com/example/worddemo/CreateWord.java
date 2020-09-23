package com.example.worddemo;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.asm.Type;
import com.example.worddemo.bean.FrameContent;
import com.example.worddemo.bean.TemplateData;
import com.example.worddemo.util.OtherUtil;
import com.example.worddemo.util.WordUtil1;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

import static com.example.worddemo.util.WordUtil.generateTOC;

/**
 * @author ys
 * @date 2020/9/21 11:46
 */
public class CreateWord {
    public void createVulnerabilityReport(String ids, HttpServletResponse response) throws InvalidFormatException, IOException{

        // 将传过来的id串分割并存入list中
        var idArrays = Arrays.asList(ids.trim().split(","));
        // 查询包装器
        // 创建文本
        var docx = new XWPFDocument();

        // 添加封面页
        WordUtil1.addHeaderPage(docx, "广东网安科技有限公司");

        // 第一段
        var firstHeaderParagraph = docx.createParagraph();
        // 插入一级标题
        WordUtil1.setLevelTitleFirst(docx, firstHeaderParagraph, "一、整改汇总");
        // 第二段
        var secondHeaderParagraph = docx.createParagraph();
        // 插入一级标题
        WordUtil1.setLevelTitleFirst(docx, secondHeaderParagraph, "二、整改摘要");

        var sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

        // 插入整改内容


        // 设置response对象
        response.setContentType("application/octet-stream; charset=utf-8");
        response.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode("广东网安科技有限公司-漏洞整改报告.docx", "UTF-8"));
        // 获取输出流
        var out=response.getOutputStream();
        // 下载
        docx.write(out);
        out.flush();
        out.close();
    }
    public static void main (String[] args) {
        ObjectMapper objectMapper = new ObjectMapper();
        var docx = new XWPFDocument();
        // 添加封面页
        WordUtil1.addHeaderPage(docx, "广东网安科技有限公司");
        // 第一段
        var firstHeaderParagraph = docx.createParagraph();
        try {
            var root = objectMapper.readTree(OtherUtil.getFileByPath("E:\\11\\newfile.json"));
            Boolean  re = root.isArray();
            Map inputContent = (Map) JSONObject.parse(root.get("input").get(0).get("content").toString());
            Map fileContent = (Map) JSONObject.parse(root.get("files").get(0).get("content").toString());
            List<TemplateData> templateDataList = JSONObject.parseArray(root.get("data").get(0).get("yj").toString(),TemplateData.class);
            Map<String, List<TemplateData>> childList = new HashMap<>(16);
            for(TemplateData data:templateDataList){
                List<TemplateData> child = JSONObject.parseArray(root.get("data").get(0).get("data").get(data.getTemplateContentId().toString()).toString(),TemplateData.class);
                childList.put(data.getTemplateContentId().toString(),child);
            }
            Map<String,List<FrameContent>> farmeContentMap = new HashMap<>(16);
            Map frame = JSONObject.parseObject(root.get("data").get(0).get("frame").toString());
            Iterator<Map.Entry<String,JSONArray>> it = frame.entrySet().iterator();
            while (it.hasNext()){
                Map.Entry<String,JSONArray> entry = it.next();
                List<FrameContent> frameContentList = JSONObject.parseArray(entry.getValue().toString(),FrameContent.class);
                for (FrameContent content: frameContentList){
                    if(root.get("input").get(0).get("content").get(content.getTemplateContentId().toString()) != null) {
                        if (content.getTemplateFrameType().equals("1") || content.getTemplateFrameType().equals("2")) {
                            String text = root.get("input").get(0).get("content").get(content.getTemplateContentId().toString()).get(content.getTemplateFrameId().toString()).asText();
                            content.setContent(text);
                        }
                        if (content.getTemplateFrameType().equals("3")) {
                            //获取文件扩展名
                            String extension = root.get("files").get(0).get("content").get("name").get(content.getTemplateContentId().toString()).get(content.getTemplateFrameId().toString()).asText();
                            String extension1 = extension.substring(extension.lastIndexOf(".")+1);
                            content.setContent(content.getTemplateContentId().toString() + "-" + content.getTemplateFrameId().toString() +"."+ extension1);
                        }
                        farmeContentMap.put(entry.getKey(), frameContentList);
                    }

                }

            }
//            Map childList =  (Map) JSONObject.parse(root.get("data").get(0).get("data").toString());

            List<Object> tree = get(templateDataList,childList,farmeContentMap,docx);
            System.out.println(tree.toString());
            WordUtil1.generateTOC(docx,"齉");
        }catch (IOException e){
            e.printStackTrace();
        }
        try (FileOutputStream fileOut = new FileOutputStream("E:\\chromeDownload\\c.docx")) {
            docx.write(fileOut);
        }catch(IOException e){
            e.printStackTrace();
        }



    }
    private static List<Object> get (List<TemplateData> root,Map<String,List<TemplateData>> childList,Map<String,List<FrameContent>> farmeContentMap,XWPFDocument docx) throws IOException{

        List<Object> objectList = new ArrayList<>();
        for (TemplateData templateData : root ) {
            Map<String,Object> mapArr = new LinkedHashMap<String, Object>();
            if(templateData.getTemplateContentSuperior().equals("0")){
                mapArr.put("id", templateData.getTemplateContentId());
                mapArr.put("template_content_number", templateData.getTemplateContentNumber());
                mapArr.put("template_content_name",templateData.getTemplateContentName());
                mapArr.put("template_content_superior", templateData.getTemplateContentSuperior());
                mapArr.put("template_content_level", templateData.getTemplateContentLevel());
                mapArr.put("content",farmeContentMap.get(templateData.getTemplateContentId().toString()));

                var firstHeaderParagraph = docx.createParagraph();
                WordUtil1.setLevelTitleFirst(docx, firstHeaderParagraph, templateData.getTemplateContentNumber()+templateData.getTemplateContentName());
                for(FrameContent content: farmeContentMap.get(templateData.getTemplateContentId().toString())){
//                    var paargraph = docx.createParagraph();
                    // 创建内容段落
                    var childContentParagraph = docx.createParagraph();
                    var childContentRun = childContentParagraph.createRun();
                    if (content.getTemplateFrameType().equals("1") || content.getTemplateFrameType().equals("2")) {
                        childContentRun.addTab();
                        childContentRun.setText(content.getContent());

                    }
                    if (content.getTemplateFrameType().equals("3")) {
                        String filePath = "E:\\11\\"+ content.getContent();
                        int pictureType = 0;
                        String prefix = filePath.substring(filePath.lastIndexOf(".")+1).toLowerCase();
                        switch (prefix) {
                            case "png":
                                pictureType = XWPFDocument.PICTURE_TYPE_PNG;
                                break;
                            case "jpg":
                            case "jpeg":
                                pictureType = XWPFDocument.PICTURE_TYPE_JPEG;
                                break;
                            case "gif":
                                pictureType = XWPFDocument.PICTURE_TYPE_GIF;
                                break;
                            default:
                                System.out.println("图片格式错误！");
                        }
                        FileInputStream is = null;
                        try {
                            is = new FileInputStream(filePath);
                            childContentRun.addCarriageReturn();
                            childContentRun.addTab();
                            childContentRun.addPicture(is, pictureType, filePath, Units.toEMU(430), Units.toEMU(300));
                            is.close();
                        } catch (FileNotFoundException | InvalidFormatException e) {
                           e.printStackTrace();
                        }

                    }
                }

                mapArr.put("children", menuChild(templateData.getTemplateContentId().toString(),childList.get(templateData.getTemplateContentId().toString()),farmeContentMap,docx));
                objectList.add(mapArr);
            }
        }
        return objectList;
    }

    private static List<?> menuChild(String id,List<TemplateData> dataList,Map<String,List<FrameContent>> farmeContentMap,XWPFDocument docx) throws IOException{
        System.out.println(dataList.toString());
        List<Object> lists = new ArrayList<Object>();
        for (TemplateData data : dataList) {
            Map<String, Object> childArray = new LinkedHashMap<String, Object>();
            if (data.getTemplateContentSuperior().equals(id)) {
                childArray.put("id", data.getTemplateContentId());
                childArray.put("template_content_number", data.getTemplateContentNumber());
                childArray.put("template_content_name",data.getTemplateContentName());
                childArray.put("template_content_superior", data.getTemplateContentSuperior());
                childArray.put("template_content_level", data.getTemplateContentLevel());
                childArray.put("content",farmeContentMap.get(data.getTemplateContentId().toString()));
                var firstHeaderParagraph = docx.createParagraph();
                WordUtil1.setLevelTitleFirst(docx, firstHeaderParagraph, data.getTemplateContentNumber()+data.getTemplateContentName());
                for(FrameContent content: farmeContentMap.get(data.getTemplateContentId().toString())){
//                    var paargraph = docx.createParagraph();
                    var childContentParagraph = docx.createParagraph();
                    var childContentRun = childContentParagraph.createRun();
                    if (content.getTemplateFrameType().equals("1") || content.getTemplateFrameType().equals("2")) {
                        // 创建内容段落

                        childContentRun.addTab();
                        childContentRun.setText(content.getContent());
                    }
                   if (content.getTemplateFrameType().equals("3")) {
                        String filePath = "E:\\11\\"+ content.getContent();
                        int pictureType = 0;
                        String prefix = filePath.substring(filePath.lastIndexOf(".")+1).toLowerCase();
                        switch (prefix) {
                            case "png":
                                pictureType = XWPFDocument.PICTURE_TYPE_PNG;
                                break;
                            case "jpg":
                            case "jpeg":
                                pictureType = XWPFDocument.PICTURE_TYPE_JPEG;
                                break;
                            case "gif":
                                pictureType = XWPFDocument.PICTURE_TYPE_GIF;
                                break;
                            default:
                                System.out.println("图片格式错误！");
                        }
                        FileInputStream is = null;
                        try {
                            is = new FileInputStream(filePath);
                            childContentRun.addCarriageReturn();
                            childContentRun.addTab();
                            childContentRun.addPicture(is, pictureType, filePath, Units.toEMU(430), Units.toEMU(300));
                            is.close();
                        } catch (FileNotFoundException | InvalidFormatException e) {
                           e.printStackTrace();
                        }

                    }
                }
                childArray.put("children", menuChild(data.getTemplateContentId().toString(),dataList,farmeContentMap,docx));
                lists.add(childArray);
            }
        }
        return lists;
    }

}
