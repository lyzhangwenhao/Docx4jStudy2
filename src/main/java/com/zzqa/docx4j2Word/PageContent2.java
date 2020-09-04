package com.zzqa.docx4j2Word;

import com.zzqa.pojo.Characteristic;
import com.zzqa.pojo.Feature;
import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.TableUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * ClassName: PageContent2
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/5 13:55
 */
public class PageContent2 {
    private ObjectFactory factory = new ObjectFactory();

    /**
     * 生成预警/报警的表格信息
     *
     * @param wpMLPackage 传入的wpMLPackage对象
     * @param mapList
     */
    public void createPageContent2(WordprocessingMLPackage wpMLPackage, List<Map<String,String>> mapList) {
        //添加标题一：项目概述
        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "2 测点安装及配置");
        if (mapList==null || mapList.size()==0){
            return;
        }
        Docx4jUtil.addParagraph(wpMLPackage, "各机组上安装有传感器"+mapList.size()+"个，各测点配置如下表所示：");
        //创建一个表格
        Tbl tbl = factory.createTbl();
        //生成表头
        createTalbeTitle(tbl);
        //生成数据
        createTableData(wpMLPackage, tbl, mapList);
        //TODO 删除输出语句
        System.out.println("PageContent2 Success......");
        //跨行合并
//        mergeCellsVertically(tbl,0, 1, 2);
    }


    /**
     * 生成表格数据
     *
     * @param wpMLPackage
     * @param tbl
     * @param mapList
     */
    private void createTableData(WordprocessingMLPackage wpMLPackage, Tbl tbl, List<Map<String,String>> mapList) {
        Tr dataTr = null;
        for (Map<String,String> map:mapList){
            dataTr = factory.createTr();
            TableUtil.addTableTc(dataTr, map.get("id"), 1000, false, "20","black",null);
            TableUtil.addTableTc(dataTr, map.get("location"), 3000, false, "20","black",null);
            TableUtil.addTableTc(dataTr, map.get("direction"), 1500, false, "20","black",null);
            TableUtil.addTableTc(dataTr, map.get("type"), 3000, false, "20","black",null);
            tbl.getContent().add(dataTr);
        }
        wpMLPackage.getMainDocumentPart().addObject(tbl);
    }

    /**
     * 为table生成表头
     *
     * @param tbl
     */
    private void createTalbeTitle(Tbl tbl) {
        //给table添加边框
        TableUtil.addBorders(tbl,"#95b3d7","4");
        Tr tr = factory.createTr();
        //单元格居中对齐
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        TblPr tblPr = tbl.getTblPr();
        tblPr.setJc(jc);
        tbl.setTblPr(tblPr);

        //表格表头
        TableUtil.addTableTc(tr, "序号", 1000, true, "20","#ffffff","#8064a2");
        TableUtil.addTableTc(tr, "测点位置", 3000, true, "20","#ffffff","#8064a2");
        TableUtil.addTableTc(tr, "方向", 1500, true, "20","#ffffff","#8064a2");
        TableUtil.addTableTc(tr, "传感器类型", 3000, true, "20","#ffffff","#8064a2");
        //将tr添加到table中
        tbl.getContent().add(tr);
    }
}
