package com.zzqa.docx4j2Word;

import com.zzqa.pojo.FaultUnitInfo;
import com.zzqa.utils.Docx4jUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.STShd;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * ClassName: PageContent6
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/3 14:54
 */
public class PageContent6 {

    private ObjectFactory factory = new ObjectFactory();

    private NumberingCreate numberingCreate;
    private long restart = 1;

    public void createPageContent6(WordprocessingMLPackage wpMLPackage, List<FaultUnitInfo> faultUnitInfoList, NumberingCreate numberingCreate) {
        if (numberingCreate == null) {
            this.numberingCreate = new NumberingCreate(wpMLPackage);
        } else {
            this.numberingCreate = numberingCreate;
        }
        if (faultUnitInfoList == null || faultUnitInfoList.size() == 0) {
            return;
        }

        int num = 1;

        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "6 故障机组详细分析");
        for (FaultUnitInfo faultUnitInfo : faultUnitInfoList) {

//            if (restart != 1) {
                restart = numberingCreate.restart(restart, 0, 1);
//            }
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", "6." + num + " " + faultUnitInfo.getUnitName());

            Docx4jUtil.addParagraph(wpMLPackage, "我们对" + faultUnitInfo.getUnitName() + "的运行情况通过远程浏览的方式进行了连续的跟踪监测。发现：");

//            numberingCreate.createNumberedParagraph(1, 0, )
            String content = faultUnitInfo.getContent();
            //发现的情况
            if (content != null && !"".equals(content)) {
                String[] split = content.split("\n");
                for (String s : split) {
                    P p = numberingCreate.createNumberedParagraph(restart, 0, s, 100);
                    PPr pPr = p.getPPr();
                    Docx4jUtil.setParagraphShdStyle(pPr, STShd.PCT_20, "#55ff55");
                    p.setPPr(pPr);
                    wpMLPackage.getMainDocumentPart().addObject(p);
                }
            }
            //结论
            restart = numberingCreate.restart(restart, 0, 1);
            Docx4jUtil.addParagraph(wpMLPackage, "结论：");
            String conclusion = faultUnitInfo.getConclusion();
            if (conclusion != null && !"".equals(conclusion)) {
                String[] split = conclusion.split("\n");
                for (String s : split) {
                    P p = numberingCreate.createNumberedParagraph(restart, 1, s, 100);
                    PPr pPr = p.getPPr();
                    Docx4jUtil.setParagraphShdStyle(pPr, STShd.PCT_20, "#55ff55");
                    wpMLPackage.getMainDocumentPart().addObject(p);
                }
            }
            //图片
            List<String[]> imageList = faultUnitInfo.getImageList();
            if (imageList != null && imageList.size() != 0) {
                for (String[] strings : imageList) {
                    if (strings==null && strings.length<2){
                        continue;
                    }
                    String title = strings[0];
                    String imagePath = strings[1];
                    File file = new File(imagePath);
                    byte[] bytes = Docx4jUtil.convertImageToByteArray(file);
                    if (bytes==null){
                        continue;
                    }
                    Docx4jUtil.addImageToPackage(wpMLPackage, bytes);
                    Docx4jUtil.addTableTitle(wpMLPackage, title);
                }

            }
            num++;

        }

    }
}
