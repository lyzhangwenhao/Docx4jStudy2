package com.zzqa.docx4j2Word;

import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.DrawChartPieUtil;
import com.zzqa.utils.TableUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.io.File;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * ClassName: PageContent
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/29 14:38
 */
public class PageContent {

    private ObjectFactory factory = new ObjectFactory();

    /**
     * 第一章节：项目描述
     */
    class PageContent1 {

        public void createPageContent(WordprocessingMLPackage wpMLPackage,
                                      String paragraContent, int normalPart, int warningPart, int alarmPart) {
            try {
                AddingAFooter addingAFooter = new AddingAFooter();
                Relationship relationship = addingAFooter.createFooterPart(wpMLPackage, "◆ 版权所有 © 2018-2020 浙江中自庆安新能源技术有限公司 &&" +
                        "◆ 我们保留本文档和信息的全部所有权利。未经明示授权，严禁复制、使用或披露给第三方。");
                addingAFooter.createFooterReference(wpMLPackage, relationship);

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
         *
         * @param imageFile
         */
        private void deleteImageFile(File imageFile) {
            if (imageFile != null && imageFile.exists()) {
                imageFile.delete();
            }
        }
    }

    /**
     * 第二章节：测点安装及配置
     */
    class PageContent2 {


        /**
         * 生成预警/报警的表格信息
         *
         * @param wpMLPackage 传入的wpMLPackage对象
         * @param mapList
         */
        public void createPageContent(WordprocessingMLPackage wpMLPackage, List<Map<String, String>> mapList) {
            //添加标题一：项目概述
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "2 测点安装及配置");
            if (mapList == null || mapList.size() == 0) {
                return;
            }
            Docx4jUtil.addParagraph(wpMLPackage, "各机组上安装有传感器" + mapList.size() + "个，各测点配置如下表所示：");
            //创建一个表格
            Tbl tbl = factory.createTbl();
            //生成表头
            createTalbeTitle(tbl);
            //生成数据
            createTableData(wpMLPackage, tbl, mapList);
            //TODO 删除输出语句
            System.out.println("PageContent2 Success......");
        }


        /**
         * 生成表格数据
         *
         * @param wpMLPackage
         * @param tbl
         * @param mapList
         */
        private void createTableData(WordprocessingMLPackage wpMLPackage, Tbl tbl, List<Map<String, String>> mapList) {
            Tr dataTr = null;
            for (Map<String, String> map : mapList) {
                dataTr = factory.createTr();
                TableUtil.addTableTc(dataTr, map.get("id"), 1000, false, "20", "black", null);
                TableUtil.addTableTc(dataTr, map.get("location"), 3000, false, "20", "black", null);
                TableUtil.addTableTc(dataTr, map.get("direction"), 1500, false, "20", "black", null);
                TableUtil.addTableTc(dataTr, map.get("type"), 3000, false, "20", "black", null);
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
            TableUtil.addBorders(tbl, "#95b3d7", "4");
            Tr tr = factory.createTr();
            //单元格居中对齐
            Jc jc = new Jc();
            jc.setVal(JcEnumeration.CENTER);
            TblPr tblPr = tbl.getTblPr();
            tblPr.setJc(jc);
            tbl.setTblPr(tblPr);

            //表格表头
            TableUtil.addTableTc(tr, "序号", 1000, true, "20", "#ffffff", "#8064a2");
            TableUtil.addTableTc(tr, "测点位置", 3000, true, "20", "#ffffff", "#8064a2");
            TableUtil.addTableTc(tr, "方向", 1500, true, "20", "#ffffff", "#8064a2");
            TableUtil.addTableTc(tr, "传感器类型", 3000, true, "20", "#ffffff", "#8064a2");
            //将tr添加到table中
            tbl.getContent().add(tr);
        }
    }

    /**
     * 第三章节：评估标准
     */
    class PageContent3 {

        public void createPageContent(WordprocessingMLPackage wpMLPackage, List<Map<String, String>> dataList) {
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "3 评估标准");
            Docx4jUtil.addParagraph(wpMLPackage, "本报告根据《VDI3834风力发电机组振动控制标准》，并结合现场机组整体运行情况对机组运行状况进行评估，各测点振动报警值如下表所示：");
            //创建一个表格
            Tbl tbl = factory.createTbl();
            //设置表格样式
            setStyle(tbl);
            //生成表头
            createTalbeTitle(tbl);
            //数据
            createTableData(tbl, dataList);

            //将表格添加到wpMlPackage中
            wpMLPackage.getMainDocumentPart().addObject(tbl);
            //TODO 删
            System.out.println("PageContent3 Success......");
        }

        private void createTableData(Tbl tbl, List<Map<String, String>> dataList) {
            if (dataList == null || dataList.size() == 0) {
                return;
            }
            for (Map<String, String> map : dataList) {
                if (map == null || map.size() == 0) {
                    continue;
                }
                Tr tr = factory.createTr();
                TableUtil.addTableTc(tr, "评估加速度(" + (map.get("单位") != null && !"".equals(map.get("单位")) ? map.get("单位") : "g") + ")", 2500, true, "22", "#000000", null);
                TableUtil.addTableTc(tr, (map.get("主轴承1") != null && !"".equals(map.get("主轴承1"))) ? map.get("主轴承1") : "\\", 1500, false, "22", "#000000", null);
                TableUtil.addTableTc(tr, (map.get("主轴承2") != null && !"".equals(map.get("主轴承2"))) ? map.get("主轴承2") : "\\", 1500, false, "22", "#000000", null);
                TableUtil.addTableTc(tr, (map.get("齿轮箱输入轴") != null && !"".equals(map.get("齿轮箱输入轴"))) ? map.get("齿轮箱输入轴") : "\\", 1500, false, "22", "#000000", null);
                TableUtil.addTableTc(tr, (map.get("齿轮箱二级齿圈") != null && !"".equals(map.get("齿轮箱二级齿圈"))) ? map.get("齿轮箱二级齿圈") : "\\", 1500, false, "22", "#000000", null);
                TableUtil.addTableTc(tr, (map.get("齿轮箱输出轴") != null && !"".equals(map.get("齿轮箱输出轴"))) ? map.get("齿轮箱输出轴") : "\\", 1500, false, "22", "#000000", null);
                TableUtil.addTableTc(tr, (map.get("发电机前轴承") != null && !"".equals(map.get("发电机前轴承"))) ? map.get("发电机前轴承") : "\\", 1500, false, "22", "#000000", null);
                TableUtil.addTableTc(tr, (map.get("发电机后轴承") != null && !"".equals(map.get("发电机后轴承"))) ? map.get("发电机后轴承") : "\\", 1500, false, "22", "#000000", null);
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

    /**
     * 第四章节：分析结论
     */
    class PageContent4 {

        public void createPageContent(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList) {
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "4 分析结论");
            Docx4jUtil.addParagraph(wpMLPackage, "评估依据：结合当前风况和CS2000振动监测系统所测得振动数据，以及时域波形图、频谱图、包络谱图、趋势图。");
            //创建一个表格
            createTable(wpMLPackage, mapTitle, dataList);
            //TODO 删
            System.out.println("PageContent7 Success......");
        }
    }

    /**
     * 第五章节：处理建议
     */
    class PageContent5 {
        private NumberingCreate numberingCreate;
        private long restart = 1;

        public void createPageContent(WordprocessingMLPackage wpMLPackage, String reportName, List<Map<String, Object>> list, NumberingCreate numberingCreate) {
            if (numberingCreate == null) {
                this.numberingCreate = new NumberingCreate(wpMLPackage);
            } else {
                this.numberingCreate = numberingCreate;
            }
            //"の"
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "5 处理建议");
            if (list == null || list.size() == 0) {
                return;
            }
            Tbl tbl = factory.createTbl();
            //设置表格样式
            setStyle(tbl);
            //创建表头
            createTblTitle(tbl,reportName);
            //添加数据
            createTblData(tbl, list);
            //将表格添加到文档中
            wpMLPackage.getMainDocumentPart().addObject(tbl);
            //TODO 删
            System.out.println("PageContent5 Success......");
        }

        /**
         * 创建表格数据
         * @param tbl
         * @param list
         */
        private void createTblData(Tbl tbl, List<Map<String, Object>> list) {
            if (list == null || list.size() == 0) {
                return;
            }
            for (Map<String, Object> map : list) {
                if (map == null || map.size() == 0 || map.get("unit_name")==null) {
                    continue;
                }
                Tr tr = factory.createTr();
                TableUtil.addTableTc(tr, (String)map.get("unit_name"), 2000, false, "20", "#000000", null);
                String[] advises = (String[]) map.get("advise");
                P[] ps = new P[advises.length];
                int index = 0;
                for (String advise : advises) {
//                    P p = numberingCreate.createNumberedParagraph(restart, 0, advise, 0);
                    P p = factory.createP();
                    Text text = factory.createText();
                    R r = factory.createR();
                    PPr pPr = factory.createPPr();
                    //添加内容到段落
                    text.setValue(advise);
                    r.getContent().add(text);
                    p.getContent().add(r);
                    //设置缩进
                    PPrBase.Ind ind = pPr.getInd();
                    if (ind==null){
                        ind = factory.createPPrBaseInd();
                    }
                    ind.setFirstLineChars(BigInteger.ZERO);
                    pPr.setInd(ind);
                    //去除段后格式
                    Docx4jUtil.setSpacing(pPr);
                    p.setPPr(pPr);
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
        private void createTblTitle(Tbl tbl, String reportName) {

            Tr tr1 = factory.createTr();
            Tr tr2 = factory.createTr();
            //第一行表头
            TableUtil.addTableTc(tr1, reportName, 2000, true, "22", "#000000", null);
            TableUtil.addTableTc(tr1, reportName, 7000, true, "22", "#000000", null);
            //第二行表头
            TableUtil.addTableTc(tr2, "机组编号", 2000, true, "22", "#000000", null);
            TableUtil.addTableTc(tr2, "处理意见", 7000, true, "22", "#000000", null);

            //将tr添加到table中
            tbl.getContent().add(tr1);
            tbl.getContent().add(tr2);
            //合并
            TableUtil.mergeCellsHorizontal(tbl, 0, 0, 1);
        }
    }

    /**
     * 第六章节：故障机组详细分析
     */
    class PageContent6 {
        private NumberingCreate numberingCreate;
        private long restart = 1;

        public void createPageContent(WordprocessingMLPackage wpMLPackage, List<Map<String, Object>> dataList, NumberingCreate numberingCreate) {
            if (numberingCreate == null) {
                this.numberingCreate = new NumberingCreate(wpMLPackage);
            } else {
                this.numberingCreate = numberingCreate;
            }
            if (dataList == null || dataList.size() == 0) {
                return;
            }

            int num = 1;

            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "6 故障机组详细分析");
            for (Map<String, Object> data : dataList) {

//            if (restart != 1) {
                restart = numberingCreate.restart(restart, 0, 1);
//            }
                if (data == null || data.get("unit_name") == null) {   //如果机组名称为空，则继续下一个
                    continue;
                }
                wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading2", "6." + num + " " + data.get("unit_name"));

                Docx4jUtil.addParagraph(wpMLPackage, "我们对" + data.get("unit_name") + "的运行情况通过远程浏览的方式进行了连续的跟踪监测。发现：");

//            numberingCreate.createNumberedParagraph(1, 0, )
                String[] contents = (String[]) data.get("contents");

                //发现的情况
                if (contents != null && contents.length > 0) {
                    for (String content : contents) {
                        P p = numberingCreate.createNumberedParagraph(restart, 0, content, 100);
                        //设置段落背景颜色
                        PPr pPr = p.getPPr();
                        Docx4jUtil.setParagraphShdStyle(pPr, STShd.PCT_20, "#55ff55");
                        p.setPPr(pPr);
                        wpMLPackage.getMainDocumentPart().addObject(p);
                    }
                }
                //结论
                restart = numberingCreate.restart(restart, 0, 1);
                Docx4jUtil.addParagraph(wpMLPackage, "结论：");
                String[] conclusions = (String[]) data.get("conclusions");
                if (conclusions != null && conclusions.length > 0) {
                    for (String conclusion : conclusions) {
                        P p = numberingCreate.createNumberedParagraph(restart, 1, conclusion, 100);
                        PPr pPr = p.getPPr();
                        Docx4jUtil.setParagraphShdStyle(pPr, STShd.PCT_20, "#55ff55");
                        wpMLPackage.getMainDocumentPart().addObject(p);
                    }
                }
                //图片
                List<String[]> imageList = (List<String[]>) data.get("imageList");
                if (imageList != null && imageList.size() != 0) {
                    for (String[] strings : imageList) {
                        if (strings == null && strings.length < 2) {
                            continue;
                        }
                        String title = strings[0];
                        String imagePath = strings[1];
                        File file = new File(imagePath);
                        byte[] bytes = Docx4jUtil.convertImageToByteArray(file);
                        if (bytes == null) {
                            continue;
                        }
                        Docx4jUtil.addImageToPackage(wpMLPackage, bytes);
                        Docx4jUtil.addTableTitle(wpMLPackage, title);
                    }

                }
                num++;
            }
            //TODO 删
            System.out.println("PageContent6 Success......");

        }
    }

    /**
     * 第七章节：正常机组数据
     */
    class PageContent7 {

        public void createPageContent(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList) {
            wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "7 正常机组数据");
            //创建一个表格
            createTable(wpMLPackage, mapTitle, dataList);
            //TODO 删
            System.out.println("PageContent7 Success......");
        }
    }

    /**
     * 第八章节：补充说明
     */
    class PageContent8 {
        private long restart = 1;
        private NumberingCreate numberingCreate;

        /**
         * 固定的第8部分内容
         *
         * @param wpMLPackage
         */
        public void createPageContent(WordprocessingMLPackage wpMLPackage, NumberingCreate numberingCreate) {
            if (numberingCreate != null) {
                this.numberingCreate = numberingCreate;
            } else {
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
         *
         * @param wpMLPackage
         * @param content
         */
        private void addParagraphNumber(WordprocessingMLPackage wpMLPackage, String content) {
            P p = numberingCreate.createNumberedParagraph(restart, 0, content, 0);
            R r = factory.createR();
            PPr pPr = p.getPPr();
            if (pPr == null) {
                pPr = factory.createPPr();
            }
            //设置行距1.5倍
            PPrBase.Spacing spacing = pPr.getSpacing();
            if (spacing == null) {
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
         *
         * @param wpMLPackage
         * @param content
         */
        private void addParagraph(WordprocessingMLPackage wpMLPackage, String content) {

            P p = factory.createP();
            R r = factory.createR();
            PPr pPr = p.getPPr();
            if (pPr == null) {
                pPr = factory.createPPr();
            }


            //设置行距1.5倍
            PPrBase.Spacing spacing = pPr.getSpacing();
            if (spacing == null) {
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

    private void createTable(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList) {
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
        //机组编号列固定不变，其他标题动态生成
        TableUtil.addTableTc(titleTr1, "机组编号", 1500, true, "22", "black", null);//该单元格固定
        TableUtil.addTableTc(titleTr2, "机组编号", 1500, true, "22", "black", null);//该单元格固定
        //合并
        TableUtil.mergeCellsVertically(tbl, 0, 0, 1);
        Set<Map.Entry<String, String[]>> entries = mapTitle.entrySet();
        //控制第一行合并起始
        int start = 1;
        //记录标题所在位置，后面插入数据时根据该位置将数据插入到对应的标题列下面
        Map<Integer, String> titlePosition = new HashMap<>();
        Integer index = 2;  //除去编号，从第2列开始
        for (Map.Entry<String, String[]> entry : entries) {
            if (entry == null) {
                continue;
            }
            String key = entry.getKey();
            String[] value = entry.getValue();
            if (value == null || value.length < 1) {
                continue;
            }
            for (String s : value) {
                //添加第一行标题
                TableUtil.addTableTc(titleTr1, key, 1500, true, "22", "black", null);
                //添加第二行标题
                TableUtil.addTableTc(titleTr2, s, 1500, true, "22", "black", null);
                //如果行标题跟第二行名称相同，则对标题进行合并
                if (key != null && key.equals(s)) {
                    TableUtil.mergeCellsVertically(tbl, index - 1, 0, 1);
                }
                //存储标题位置
                titlePosition.put(index, s);
                index++;
            }
            //合并
            TableUtil.mergeCellsHorizontal(tbl, 0, start, start + value.length);
            start += value.length;
        }
        //下面插入数据部分
        if (dataList == null || dataList.size() == 0) {
            return;
        }
        for (Map<String, String> map : dataList) {
            if (map == null || map.size() == 0) {
                continue;
            }
            Tr tr = factory.createTr();
            TableUtil.addTableTc(tr, (map.get("机组编号") != null && !"".equals(map.get("机组编号")) ? map.get("机组编号") : "\\"), 2500, true, "22", "#000000", null);
            for (int i = 0; i < titlePosition.size(); i++) {//根据位置插入相应的数据
                TableUtil.addTableTc(tr, (map.get(titlePosition.get(i + 2)) != null && !"".equals(map.get(titlePosition.get(i + 2)))) ? map.get(titlePosition.get(i + 2)) : "\\", 1500, false, "22", "#000000", null);
            }
            tbl.getContent().add(tr);
        }
    }

    /**
     * 设置表格样式
     *
     * @param tbl
     */
    private void setStyle(Tbl tbl) {
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
