package com.zzqa.docx4j2Word;

import com.zzqa.utils.TableUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.util.List;

/**
 * ClassName: PageContent5
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/4 13:40
 */
public class PageContent5 {
    private ObjectFactory factory = new ObjectFactory();
    private NumberingCreate numberingCreate;
    private long restart = 1;

    public void createPageContent(WordprocessingMLPackage wpMLPackage, String name, List<String> list, NumberingCreate numberingCreate) {
        if (numberingCreate==null){
            this.numberingCreate = new NumberingCreate(wpMLPackage);
        }else {
            this.numberingCreate = numberingCreate;
        }
        //"の"
        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "5 处理建议");
        if (list == null || list.size() == 0) {
            return;
        }
        Tbl tbl = factory.createTbl();
        //给table添加边框
        TableUtil.addBorders(tbl, "#4bacc6", "4");
        //单元格居中对齐
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        TblPr tblPr = tbl.getTblPr();
        tblPr.setJc(jc);
        tbl.setTblPr(tblPr);
        //创建表头
        createTblTitle(tbl);
        //添加数据
        createTblData(tbl, list);
        //将表格添加到文档中
        wpMLPackage.getMainDocumentPart().addObject(tbl);
        //TODO 删
        System.out.println("PageContent5 Success......");
    }

    private void createTblData(Tbl tbl, List<String> list) {
        if (list == null || list.size() == 0) {
            return;
        }
        Tr tr = null;
        for (String line : list) {
            if (line == null || line.length() < 2) {
                continue;
            }
            String[] split = line.split("の");
            tr = factory.createTr();
            TableUtil.addTableTc(tr, split[0], 2000, false, "20", "#000000", null);
            String[] advise = split[1].split("\n");
            P[] ps = new P[advise.length];
            int index = 0;
            for (String s : advise) {
                P p = numberingCreate.createNumberedParagraph(restart, 0, s, 0);
                ps[index] = p;
                index++;
            }
            TableUtil.addTableTcNumbering(tr, ps, 7000, false, "20", "#000000", null);
            restart = numberingCreate.restart(restart, 0, 1);
            tbl.getContent().add(tr);
        }
    }


    /**
     * 创建表头
     *
     * @param tbl
     */
    private void createTblTitle(Tbl tbl) {

        Tr tr1 = factory.createTr();
        Tr tr2 = factory.createTr();
        //第一行表头
        TableUtil.addTableTc(tr1, "上海东滩风电场", 2000, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "上海东滩风电场", 7000, true, "22", "#ffffff", "#4bacc6");
        //第二行表头
        TableUtil.addTableTc(tr2, "机组编号", 2000, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "处理意见", 7000, true, "22", "#ffffff", "#4bacc6");

        //将tr添加到table中
        tbl.getContent().add(tr1);
        tbl.getContent().add(tr2);
        //合并
        TableUtil.mergeCellsHorizontal(tbl, 0, 0, 1);
    }
}
