package com.zzqa.docx4j2Word;

import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.TableUtil;
import org.apache.log4j.Logger;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.util.List;
import java.util.Map;


/**
 * ClassName: PageContent4
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/7 10:41
 */
public class PageContent4 {
    private ObjectFactory factory = new ObjectFactory();

    public void createPageContent(WordprocessingMLPackage wpMLPackage){
        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "4 分析结论");
        Docx4jUtil.addParagraph(wpMLPackage, "评估依据：结合当前风况和CS2000振动监测系统所测得振动数据，以及时域波形图、频谱图、包络谱图、趋势图。");
        //创建一个表格
        Tbl tbl = factory.createTbl();
        //给table添加边框
        TableUtil.addBorders(tbl, "#4bacc6", "4");
        //单元格居中对齐
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        TblPr tblPr = tbl.getTblPr();
        tblPr.setJc(jc);
        tbl.setTblPr(tblPr);
        //生成表头
        createTalbeTitle(tbl);
        //数据
//        createTableData(tbl, dataList);

        //将表格添加到wpMlPackage中
        wpMLPackage.getMainDocumentPart().addObject(tbl);
        //TODO 删
        System.out.println("PageContent3 Success......");
    }

    private void createTable(Tbl tbl,Map<String,String[]> mapTitle){

    }

    private void createTableData(Tbl tbl, List<Map<String, String>> dataList){
        if (dataList==null||dataList.size()==0){
            return;
        }
        for (Map<String, String> map : dataList) {
            if (map==null||map.size()==0){
                continue;
            }
            Tr tr = factory.createTr();
            TableUtil.addTableTc(tr, "评估加速度("+(map.get("单位")!=null&&"".equals(map.get("单位"))?map.get("单位"):"g")+")", 2500, true, "22", "#000000", null);
            TableUtil.addTableTc(tr, (map.get("主轴承1")!=null&&!"".equals(map.get("主轴承1")))?map.get("主轴承1"):"\\", 1500, false, "22", "#000000", null);
            TableUtil.addTableTc(tr, (map.get("主轴承2")!=null&&!"".equals(map.get("主轴承2")))?map.get("主轴承2"):"\\", 1500, false, "22", "#000000", null);
            TableUtil.addTableTc(tr, (map.get("齿轮箱输入轴")!=null&&!"".equals(map.get("齿轮箱输入轴")))?map.get("齿轮箱输入轴"):"\\", 1500, false, "22", "#000000", null);
            TableUtil.addTableTc(tr, (map.get("齿轮箱二级齿圈")!=null&&!"".equals(map.get("齿轮箱二级齿圈")))?map.get("齿轮箱二级齿圈"):"\\", 1500, false, "22", "#000000", null);
            TableUtil.addTableTc(tr, (map.get("齿轮箱输出轴")!=null&&!"".equals(map.get("齿轮箱输出轴")))?map.get("齿轮箱输出轴"):"\\", 1500, false, "22", "#000000", null);
            TableUtil.addTableTc(tr, (map.get("发电机前轴承")!=null&&!"".equals(map.get("发电机前轴承")))?map.get("发电机前轴承"):"\\", 1500, false, "22", "#000000", null);
            TableUtil.addTableTc(tr, (map.get("发电机后轴承")!=null&&!"".equals(map.get("发电机后轴承")))?map.get("发电机后轴承"):"\\", 1500, false, "22", "#000000", null);
            tbl.getContent().add(tr);
        }
    }

    /**
     * 为table生成表头
     *
     * @param tbl
     */
    private void createTalbeTitle(Tbl tbl) {

        Tr tr1 = factory.createTr();    //第一行表头
        Tr tr2 = factory.createTr();    //第二行表头
        //表格第一行表头
        TableUtil.addTableTc(tr1, "组件", 2500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "主轴承", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "主轴承", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "齿轮箱", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "齿轮箱", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "齿轮箱", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "发电机", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr1, "发电机", 1500, true, "22", "#ffffff", "#4bacc6");
        //将tr添加到table中
        tbl.getContent().add(tr1);
        //表格第二行表头
        TableUtil.addTableTc(tr2, "组件", 2500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "主轴承1", 1500, true, "21", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "主轴承2", 1500, true, "21", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "齿轮箱输入轴", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "齿轮箱二级齿圈", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "齿轮箱输入轴", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "发电机前轴承", 1500, true, "22", "#ffffff", "#4bacc6");
        TableUtil.addTableTc(tr2, "发电机后轴承", 1500, true, "22", "#ffffff", "#4bacc6");
        //将tr添加到table中
        tbl.getContent().add(tr2);
        //合并单元格
        TableUtil.mergeCellsVertically(tbl, 0, 0, 1);   //组件
        TableUtil.mergeCellsHorizontal(tbl, 0, 1, 2);   //主轴承
        TableUtil.mergeCellsHorizontal(tbl, 0, 3, 5);   //齿轮箱
        TableUtil.mergeCellsHorizontal(tbl, 0, 6, 7);   //发电机
    }
}