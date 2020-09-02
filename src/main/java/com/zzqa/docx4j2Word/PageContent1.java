package com.zzqa.docx4j2Word;

import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.DrawChartPieUtil;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;

/**
 * ClassName: PageContent
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/3 11:19
 */
public class PageContent1 {
    private ObjectFactory objectFactory = new ObjectFactory();

    public void createPageContent1(WordprocessingMLPackage wpMLPackage,
                                                      String paragraContent, int normalPart, int warningPart, int alarmPart){
        try {
            AddingAFooter addingAFooter = new AddingAFooter();
            Relationship relationship = addingAFooter.createFooterPart(wpMLPackage,"◆ 版权所有 © 2018-2020 浙江中自庆安新能源技术有限公司 &&" +
                    "◆ 我们保留本文档和信息的全部所有权利。未经明示授权，严禁复制、使用或披露给第三方。");
            addingAFooter.createFooterReference(wpMLPackage,relationship);

            AddingAHeader addingAHeader = new AddingAHeader();
            Relationship headerPart = addingAHeader.createHeaderPart(wpMLPackage, "咨询电话：4000093668-7 &&" + "网址：www.windit.com.cn ");
            addingAHeader.createHeaderReference(wpMLPackage, headerPart);

            //添加标题一：项目概述
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "1 项目概述");
            //文本内容
            Docx4jUtil.addParagraph(wpMLPackage, paragraContent);

            //根据数据生成饼状图
            File pieImage = DrawChartPieUtil.getImageFile("机组运行状况", normalPart, warningPart, alarmPart);
            byte[] pieImageBytes = Docx4jUtil.convertImageToByteArray(pieImage);
            deleteImageFile(pieImage);
            Docx4jUtil.addImageToPackage(wpMLPackage, pieImageBytes);

            //TODO 删除输出语句
            System.out.println("PageContent1 Success......");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    /**
     * 删除已经添加进文档中的图片
     * @param imageFile
     */
    private void deleteImageFile(File imageFile){
        if (imageFile != null && imageFile.exists()){
            imageFile.delete();
        }
    }
}
