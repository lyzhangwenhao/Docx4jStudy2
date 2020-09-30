package com.zzqa.test;

import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.TableUtil;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;

/**
 * ClassName: ParagraphTest
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/7 10:50
 */
public class ParagraphTest {
    private static ObjectFactory objectFactory = new ObjectFactory();
    public static void main(String[] args) {
        try {
            WordprocessingMLPackage wpMLPackage = WordprocessingMLPackage.createPackage();

            Tbl tbl = objectFactory.createTbl();
            Tr tr = objectFactory.createTr();
            Tc tc = objectFactory.createTc();
            TableUtil.addBorders(tbl,"#95b3d7","4");
//            TableUtil.addTableTc(tr, "正常：运行状态处于正常状态，机组可照常运行；", 8000, false, "20", "black", "#00ff00");
            tbl.getContent().add(tr);

            P p = objectFactory.createP();
            R r = objectFactory.createR();
            PPr pPr = p.getPPr();
            if (pPr==null){
                pPr = objectFactory.createPPr();
            }
            //缩进2字符
            PPrBase.Ind ind = pPr.getInd();
            if (ind==null){
                ind = objectFactory.createPPrBaseInd();
            }
            ind.setFirstLineChars(BigInteger.valueOf(200));
            pPr.setInd(ind);
            //设置行距1.5倍
            PPrBase.Spacing spacing = pPr.getSpacing();
            if (spacing==null){
                spacing = objectFactory.createPPrBaseSpacing();
            }
            spacing.setLineRule(STLineSpacingRule.AUTO);
            spacing.setLine(BigInteger.valueOf(360));
            //段前段后
//        spacing.setBeforeLines(BigInteger.valueOf(50));
//        spacing.setAfterLines(BigInteger.valueOf(50));

            Text text = objectFactory.createText();
            text.setValue("正常：运行状态处于正常状态，机组可照常运行；");
            r.getContent().add(text);

            RPr rPr = objectFactory.createRPr();

            r.setRPr(rPr);
            pPr.setSpacing(spacing);
            p.getContent().add(r);
            p.setPPr(pPr);
            //设置背景颜色
            CTShd shd = new CTShd();
            shd.setVal(STShd.CLEAR);
            shd.setColor("auto");
            shd.setFill("#00ff00");
            rPr.setShd(shd);


            P p1 = objectFactory.createP();
            Text text1 = objectFactory.createText();
            text1.setValue("注意：机组存在早期故障特征，可照常运行，应关注机组运行状况，加强日常检查和维护；");
            R r1 = objectFactory.createR();
            r1.getContent().add(text1);
            PPr pPr1 = objectFactory.createPPr();
            p1.setPPr(pPr1);
            p1.getContent().add(r1);
            Docx4jUtil.setParagraphShdStyle(pPr1,STShd.SOLID,"#00ff00");

            P[] ps = {p,p1};
            TableUtil.addTableTcNumbering(tr,ps,8000 , false, "20", "#00ff00", null);


            wpMLPackage.getMainDocumentPart().addObject(tbl);

            wpMLPackage.getMainDocumentPart().addObject(p1);

            wpMLPackage.getMainDocumentPart().addObject(p);


            wpMLPackage.save(new File("D:/AutoExport/docx4j2/test/paragraphTest.docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 添加TableCell
     *
     * @param tableRow
     * @param content
     */
    public static void addTableTc(Tr tableRow, String content, int width, boolean isBold, String fontSize, String fontColor, String backgroundcolor) {
        Tc tc = objectFactory.createTc();
        P p = objectFactory.createP();
        R r = objectFactory.createR();
        RPr rPr = objectFactory.createRPr();
        Text text = objectFactory.createText();
        //禁止行号(不设置没什么影响)
        BooleanDefaultTrue bCs = rPr.getBCs();
        if (bCs == null) {
            bCs = new BooleanDefaultTrue();
        }
        bCs.setVal(true);
        rPr.setBCs(bCs);

        //设置宽度
        setCellWidth(tc, width);
        //生成段落添加到单元格中
        text.setValue(content);
        //设置字体颜色，加粗
        Docx4jUtil.setFontColor(rPr, isBold, fontColor);
        //设置字体
        Docx4jUtil.setFont(rPr, "宋体");
        //设置字体大小
        Docx4jUtil.setFontSize(rPr, fontSize);
        //将样式添加到段落中
        r.getContent().add(rPr);

        r.getContent().add(text);
        p.getContent().add(r);
        tc.getContent().add(p);
        //去除段后格式
        PPr pPr = p.getPPr();
        if (pPr == null) {
            pPr = objectFactory.createPPr();
        }
        Docx4jUtil.setSpacing(pPr);
        p.setPPr(pPr);

        if (backgroundcolor != null && !"".equals(backgroundcolor)) {
            //设置背景颜色
            CTShd shd = new CTShd();
            shd.setVal(STShd.CLEAR);
            shd.setColor("auto");
            shd.setFill(backgroundcolor);
            tc.getTcPr().setShd(shd);
        }

        tableRow.getContent().add(tc);
    }

    /**
     * 本方法创建一个单元格属性集对象和一个表格宽度对象. 将给定的宽度设置到宽度对象然后将其添加到
     * 属性集对象. 最后将属性集对象设置到单元格中.
     */
    private static void setCellWidth(Tc tableCell, int width) {
        TcPr tableCellProperties = new TcPr();
        TblWidth tableWidth = new TblWidth();
        tableWidth.setW(BigInteger.valueOf(width));
        tableCellProperties.setTcW(tableWidth);
        tableCell.setTcPr(tableCellProperties);
    }

}
