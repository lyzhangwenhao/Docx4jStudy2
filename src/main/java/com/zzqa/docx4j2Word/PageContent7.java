package com.zzqa.docx4j2Word;

import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.TableUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * ClassName: PageContent7
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/28 16:39
 */
public class PageContent7 {
    private ObjectFactory factory = new ObjectFactory();

    public void createPageContent(WordprocessingMLPackage wpMLPackage, Map<String,String[]> mapTitle) {
        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "7 正常机组数据");
        //创建一个表格
        createTable(wpMLPackage, mapTitle);

        //TODO 删
        System.out.println("PageContent7 Success......");
    }

    private void createTable(WordprocessingMLPackage wpMLPackage,Map<String,String[]> mapTitle){
        Tbl tbl = factory.createTbl();
        //将表格添加到wpMlPackage中
        wpMLPackage.getMainDocumentPart().addObject(tbl);
        //设置样式
        setStyle(tbl);
        //生成表头
        Tr titleTr1 = factory.createTr();
        Tr titleTr2 = factory.createTr();
        tbl.getContent().add(titleTr1);
        tbl.getContent().add(titleTr2);
        TableUtil.addTableTc(titleTr1, "机组编号", 1500, true, "22", "black", null);//该单元格固定
        TableUtil.addTableTc(titleTr2, "机组编号", 1500, true, "22", "black", null);//该单元格固定
        //合并
        TableUtil.mergeCellsVertically(tbl, 0, 0, 1);
        Set<Map.Entry<String, String[]>> entries = mapTitle.entrySet();
        //控制第一行合并起始
        int start = 1;
        //记录标题所在位置
        Map<Integer,String> titlePosition = new HashMap<>();
        Integer index = 2;  //除去编号，从第2列开始
        for (Map.Entry<String, String[]> entry : entries) {
            if (entry==null){
                continue;
            }
            String key = entry.getKey();
            String[] value = entry.getValue();
            if (value==null||value.length<1){
                continue;
            }
            for (String s : value) {
                //添加第一行标题
                TableUtil.addTableTc(titleTr1, key, 1500, true, "22", "black", null);
                //添加第二行标题
                TableUtil.addTableTc(titleTr2, s, 1500, true, "22", "black", null);
                //存储标题位置
                titlePosition.put(index, s);
                index++;
            }
            //合并
            TableUtil.mergeCellsHorizontal(tbl, 0, start, start+value.length);
            start+=value.length;
        }
        //TODO 删
        System.out.println(titlePosition);
    }

    /**
     * 设置表格样式
     * @param tbl
     */
    private void setStyle(Tbl tbl){
        //给table添加边框
        TableUtil.addBorders(tbl, "#4bacc6", "4");
        //单元格居中对齐
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        TblPr tblPr = tbl.getTblPr();
        tblPr.setJc(jc);
        tbl.setTblPr(tblPr);
    }
}
