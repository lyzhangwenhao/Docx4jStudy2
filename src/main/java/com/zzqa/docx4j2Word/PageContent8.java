package com.zzqa.docx4j2Word;

import com.zzqa.utils.Docx4jUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.math.BigInteger;

/**
 * ClassName: PageContent8
 * Description:
 *
 * @author 张文豪
 * @date 2020/8/6 14:56
 */
public class PageContent8 {
    private ObjectFactory factory = new ObjectFactory();

    private long restart =1;
    private NumberingCreate numberingCreate;
    /**
     * 固定的第8部分内容
     * @param wpMLPackage
     */
    public void createPageContent8(WordprocessingMLPackage wpMLPackage,NumberingCreate numberingCreate){
        if (numberingCreate!=null){
            this.numberingCreate = numberingCreate;
        }else {
            numberingCreate = new NumberingCreate(wpMLPackage);
        }

        try {

            Docx4jUtil.addNextPage(wpMLPackage);
            restart = numberingCreate.restart(1, 0, 1);
            //添加标题四：补充说明
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "8 补充说明");
            addParagraphNumber(wpMLPackage, "本报告涂改无效。");
            addParagraphNumber(wpMLPackage, "需要委托方提供机组详细的传动链参数（主轴承参数、齿轮箱参数、发电机参数、偏航系统参数）以保证报告的准确性。");
            addParagraphNumber(wpMLPackage, "未经本中心书面许可，部分复制、摘用或篡改本报告内容，引起法律纠纷，责任自负。");
            addParagraphNumber(wpMLPackage, "本检测报告是基于对机组所安装的CS2000系统的振动数据所获得的信息而编制的，因此，本报告对机组状态所做分析仅供参考。浙江中自庆安新能源技术有限公司给出的所有信息、忠告和建议都仅是基于我们的观察、分析和经验。对于设备状况的最终判断以及所需采取的维护措施，由用户自行决定。");
            addParagraphNumber(wpMLPackage, "对检测报告若有异议，请于收到报告之日起一个月内向本中心提出，逾期不再受理。");

            Docx4jUtil.addBr(wpMLPackage, 18);
            addParagraph(wpMLPackage, "地址：杭州经济技术开发区6号路260号中自科技园");
            addParagraph(wpMLPackage, "邮编：310018");
            addParagraph(wpMLPackage, "电话：0571-28995840");
            addParagraph(wpMLPackage, "传真：0571-28995841");
            //TODO 删除输出语句
            System.out.println("PageContent8 Success......");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    /**
     * 添加正文段落(列表)
     * @param wpMLPackage
     * @param content
     */
    private void addParagraphNumber(WordprocessingMLPackage wpMLPackage,String content){
        P p = numberingCreate.createNumberedParagraph(restart, 0, content, 0);
        R r = factory.createR();
        PPr pPr = p.getPPr();
        if (pPr==null){
            pPr = factory.createPPr();
        }
        //设置行距1.5倍
        PPrBase.Spacing spacing = pPr.getSpacing();
        if (spacing==null){
            spacing = factory.createPPrBaseSpacing();
        }
        spacing.setLineRule(STLineSpacingRule.AUTO);
        spacing.setLine(BigInteger.valueOf(360));
        spacing.setAfter(new BigInteger("0"));
        //段前段后
//        spacing.setBeforeLines(BigInteger.valueOf(50));
//        spacing.setAfterLines(BigInteger.valueOf(50));
        pPr.setSpacing(spacing);
        p.getContent().add(r);
        p.setPPr(pPr);
        wpMLPackage.getMainDocumentPart().addObject(p);
    }
    /**
     * 添加正文段落
     * @param wpMLPackage
     * @param content
     */
    private void addParagraph(WordprocessingMLPackage wpMLPackage,String content){

        P p = factory.createP();
        R r = factory.createR();
        PPr pPr = p.getPPr();
        if (pPr==null){
            pPr = factory.createPPr();
        }


        //设置行距1.5倍
        PPrBase.Spacing spacing = pPr.getSpacing();
        if (spacing==null){
            spacing = factory.createPPrBaseSpacing();
        }
        spacing.setLineRule(STLineSpacingRule.AUTO);
        spacing.setLine(BigInteger.valueOf(360));
        spacing.setAfter(new BigInteger("0"));
        //段前段后
//        spacing.setBeforeLines(BigInteger.valueOf(50));
//        spacing.setAfterLines(BigInteger.valueOf(50));

        Text text = factory.createText();
        text.setValue(content);
        r.getContent().add(text);

        pPr.setSpacing(spacing);
        p.getContent().add(r);
        p.setPPr(pPr);
        wpMLPackage.getMainDocumentPart().addObject(p);
    }
}
