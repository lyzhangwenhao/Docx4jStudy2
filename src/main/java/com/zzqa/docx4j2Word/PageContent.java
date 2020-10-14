package com.zzqa.docx4j2Word;

import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.DrawChartPieUtil;
import com.zzqa.utils.TableUtil;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.math.BigInteger;
import java.util.*;

/**
 * ClassName: PageContent
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/29 14:38
 */
public class PageContent {

    private ObjectFactory factory = new ObjectFactory();

    private Logger logger = LoggerFactory.getLogger(PageContent.class);
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
                logger.error("PageContent1生成异常");
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
            try {
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
            } catch (Exception e) {
                e.printStackTrace();
                logger.error("PageContent2生成异常");
            }
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
                TableUtil.addTableTc(dataTr, map.get("ch"), 1000, false, "22", "black", null);
                TableUtil.addTableTc(dataTr, map.get("name"), 3000, false, "22", "black", null);
                TableUtil.addTableTc(dataTr, map.get("direction"), 1500, false, "22", "black", null);
                TableUtil.addTableTc(dataTr, map.get("type"), 3000, false, "22", "black", null);
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

        public void createPageContent(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList) {
            try {


                wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "3 评估标准");
                Docx4jUtil.addParagraph(wpMLPackage, "本报告根据《VDI3834风力发电机组振动控制标准》，并结合现场机组整体运行情况对机组运行状况进行评估，各测点振动报警值如下表所示：");
                //创建一个表格
                Tbl tbl = factory.createTbl();
                //设置表格样式
                setStyle(tbl);

                createTable(wpMLPackage, mapTitle, dataList, 2200);

                //将表格添加到wpMlPackage中
                wpMLPackage.getMainDocumentPart().addObject(tbl);
                //TODO 删
                System.out.println("PageContent3 Success......");
            } catch (Exception e) {
                e.printStackTrace();
                logger.error("PageContent3生成异常");
            }
        }

        private void createTable(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList, int numberIdWidth) {
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
            //设置表格宽度
//            TableUtil.setTableWidth(tbl, ""+numberIdWidth+4000);
            String color = "#4bacc6";
            //机组编号列固定不变，其他标题动态生成
            TableUtil.addTableTc(titleTr1, "组件", numberIdWidth, true, "22", "white", color);//该单元格固定
            TableUtil.addTableTc(titleTr2, "组件", numberIdWidth, true, "22", "white", color);//该单元格固定
            //合并
            TableUtil.mergeCellsVertically(tbl, 0, 0, 1);
            Set<Map.Entry<String, String[]>> entries = mapTitle.entrySet();
            //控制第一行合并起始
            int start = 1;
            //记录表格列数
            int rowNum = 0;
            for (Map.Entry<String, String[]> entry : entries) {
                if (entry == null) {
                    continue;
                }
                String[] value = entry.getValue();
                if (value == null) {
                    continue;
                }
                rowNum += value.length;
            }
            //单元格宽度
            int width = rowNum == 0 ? -1 : 10000 / rowNum;

            System.out.println(width);  // TODO
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
                    TableUtil.addTableTc(titleTr1, key, width, true, "22", "white", color);
                    //添加第二行标题
                    TableUtil.addTableTc(titleTr2, s, width, true, "22", "white", color);
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
                TableUtil.addTableTc(tr, (map.get("组件") != null && !"".equals(map.get("组件")) ? map.get("组件") : "\\"), numberIdWidth, true, "22", "#000000", null);
                for (int i = 0; i < titlePosition.size(); i++) {//根据位置插入相应的数据
                    TableUtil.addTableTc(tr, (map.get(titlePosition.get(i + 2)) != null && !"".equals(map.get(titlePosition.get(i + 2)))) ? map.get(titlePosition.get(i + 2)) : "\\", width, false, "22", "#000000", null);
                }
                tbl.getContent().add(tr);
            }
        }
    }

    /**
     * 第四章节：分析结论
     */
    class PageContent4 {

        public void createPageContent(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList) {
            try {
                wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "4 分析结论");
                Docx4jUtil.addParagraph(wpMLPackage, "评估依据：结合当前风况和CS2000振动监测系统所测得振动数据，以及时域波形图、频谱图、包络谱图、趋势图。");
                //创建一个表格
                createTable(wpMLPackage, mapTitle, dataList, 1500);
                //TODO 删
                System.out.println("PageContent4 Success......");
            } catch (Exception e) {
                e.printStackTrace();
                logger.error("PageContent4生成异常");
            }
        }
        private void createTable(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList ,int numberIdWidth) {

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
            TableUtil.addTableTc(titleTr1, "机组编号", numberIdWidth, true, "22", "black", "#4bacc6");//该单元格固定
            TableUtil.addTableTc(titleTr2, "机组编号", numberIdWidth, true, "22", "black", "#4bacc6");//该单元格固定
            //合并
            TableUtil.mergeCellsVertically(tbl, 0, 0, 1);
            //控制第一行合并起始
            int start = 1;
            //记录标题所在位置，后面插入数据时根据该位置将数据插入到对应的标题列下面
            Map<Integer, String> titlePosition = new HashMap<>();
            Integer index = 2;  //除去编号，从第2列开始
            //排序
            List<Map.Entry<String, String[]>> list = new ArrayList<>(mapTitle.entrySet());
            Collections.sort(list, (o1, o2) -> o2.getKey().compareTo(o1.getKey()));
            for (Map.Entry<String, String[]> entry : list) {
                if (entry == null) {
                    continue;
                }
                String key = entry.getKey();
                String[] value = entry.getValue();
                if (value == null || value.length < 1) {
                    continue;
                }
                //控制单元格宽度
                int width = 500;
                if ("分析结论".equals(key)){
                    width = 4500;
                }
                if ("与上季度振动趋势对比".equals(key)) {
                    width = 1000;
                }
                for (String s : value) {
                    //添加第一行标题
                    TableUtil.addTableTc(titleTr1, key, width, true, "22", "black", "#4bacc6");
                    //添加第二行标题
                    TableUtil.addTableTc(titleTr2, s, width, true, "22", "black", "#4bacc6");
                    //如果行标题跟第二行名称相同且第二行标题只有一个，则对标题进行合并
                    if (key != null && key.equals(s) && value.length==1) {
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
            int widthSum = numberIdWidth;
            for (Map<String, String> map : dataList) {
                if (map == null || map.size() == 0) {
                    continue;
                }
                Tr tr = factory.createTr();
                TableUtil.addTableTc(tr, (map.get("机组编号") != null && !"".equals(map.get("机组编号")) ? map.get("机组编号") : "\\"), numberIdWidth, true, "22", "#000000", null);
                for (int i = 0; i < titlePosition.size(); i++) {//根据位置插入相应的数据
                    String titleName = titlePosition.get(i + 2);
                    int width = 500;
                    String fontSize = "22";
                    if ("分析结论".equals(titleName)){
                        width = 4500;
                        fontSize = "21";
                    }
                    if ("与上季度振动趋势对比".equals(titleName)) {
                        width = 1000;
                    }
                    String content = map.get(titleName);
                    String backgroundColor = null;
                    if ("正常".equals(content)){
                        backgroundColor = "#00ff00";
                    }else if ("注意".equals(content)){
                        backgroundColor = "#00ffff";
                    }else if ("警告".equals(content)){
                        backgroundColor = "#ffff00";
                    }else if ("报警".equals(content)){
                        backgroundColor = "#ff00ff";
                    }else if ("危险".equals(content)){
                        backgroundColor = "#c60000";
                    }
                    widthSum += width;
                    TableUtil.addTableTc(tr, (content != null && !"".equals(map.get(titlePosition.get(i + 2)))) ? content : "\\", width, false, fontSize, "#000000", backgroundColor);
                }
                tbl.getContent().add(tr);
            }
            //在最后一行添加故障说明
            Tr tr = factory.createTr();
            Tc tc = factory.createTc();

            TableUtil.addP2Tc(tc, "备注：故障等级说明", 0, false, "22", "#000000", null,false);
            TableUtil.addP2Tc(tc, "正常：", 0, false, "22", "#000000", "#00ff00",false);
            TableUtil.addP2Tc(tc, "运行状态处于正常状态，机组可照常运行；", 0, false, "22", "#000000", null, true);
            TableUtil.addP2Tc(tc, "注意：", 0, false, "22", "#000000", "#00ffff",false);
            TableUtil.addP2Tc(tc, "机组存在早期故障特征，可照常运行，应关注机组运行状况，加强日常检查和维护；", 0, false, "22", "#000000", null, true);
            TableUtil.addP2Tc(tc, "警告：", 0, false, "22", "#000000", "#ffff00",false);
            TableUtil.addP2Tc(tc, "机组存在较明显的故障特征，机组可继续运行，现场维护人员需在2个月内检查确认故障，择机进行维修措施；", 0, false, "22", "#000000", null, true);
            TableUtil.addP2Tc(tc, "报警：", 0, false, "22", "#000000", "#ff00ff",false);
            TableUtil.addP2Tc(tc, "机组故障特征明显，故障处于劣化期，现场维护人员需在2周内检查确认故障，择机进行维修措施；", 0, false, "22", "#000000", null, true);
            TableUtil.addP2Tc(tc, "危险：", 0, false, "22", "#000000", "#c60000",false);
            TableUtil.addP2Tc(tc, "机组故障严重，须立即停机进行检查，采取维修措施。", widthSum, false, "22", "#000000", null, true);

//            TableUtil.addTableTc(tr, "", 1500, false, "22", "#000000", null);
            tr.getContent().add(tc);
            for (int i=0; i<titlePosition.size(); i++){
                TableUtil.addTableTc(tr, "", 0, false, "22", "#000000", null);
            }
            tbl.getContent().add(tr);
            TableUtil.mergeCellsHorizontal(tbl, dataList.size()+2, 0, titlePosition.size()+1);
        }

    }

    /**
     * 第五章节：处理建议
     */
    class PageContent5 {
        private NumberingCreate numberingCreate;
        private long restart = 1;

        public void createPageContent(WordprocessingMLPackage wpMLPackage, String reportName, Map<String, String[]> data, NumberingCreate numberingCreate) {
            try {
                if (numberingCreate == null) {
                    this.numberingCreate = new NumberingCreate(wpMLPackage);
                } else {
                    this.numberingCreate = numberingCreate;
                }
                wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "5 处理建议");
                if (data == null || data.size() == 0) {
                    return;
                }
                Tbl tbl = factory.createTbl();
                //设置表格样式
                setStyle(tbl);
                //创建表头
                createTblTitle(tbl, reportName);
                //添加数据
                createTblData(tbl, data);
                //将表格添加到文档中
                wpMLPackage.getMainDocumentPart().addObject(tbl);
                //TODO 删
                System.out.println("PageContent5 Success......");
            } catch (Exception e) {
                e.printStackTrace();
                logger.error("PageContent5生成异常");
            }
        }

        /**
         * 创建表格数据
         * @param tbl
         * @param data
         */
        private void createTblData(Tbl tbl, Map<String, String[]> data) {
            if (data == null || data.size() == 0) {
                return;
            }
            ArrayList<Map.Entry<String, String[]>> list = new ArrayList<>(data.entrySet());
            Collections.sort(list, Comparator.comparing(Map.Entry::getKey));
            for (Map.Entry<String, String[]> entry : list) {
                if (entry == null) {
                    continue;
                }
                String key = entry.getKey();
                String[] value = entry.getValue();
                if (key == null || value == null) {
                    continue;
                }
                Tr tr = factory.createTr();
                TableUtil.addTableTc(tr, key, 2000, false, "20", "#000000", null);
                P[] ps = new P[value.length];
                int index = 0;
                for (String advise : value) {
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
            try {
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
                int index = 0;
                for (Map<String, Object> data : dataList) {
                    index++;
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
                            int length = content.length();
                            if (content == null || length == 0) {
                                continue;
                            }
                            //定义要传的字符串
                            String parameter = null;
                            //操作（图x）格式为（图x-y）
                            int startIndex = content.lastIndexOf("（图") != -1 ? content.lastIndexOf("（图") : content.lastIndexOf("(图");
                            int endIndex = content.lastIndexOf("）") != -1 ? content.lastIndexOf("）") : content.lastIndexOf(")");
                            if (startIndex == -1 || endIndex == -1 || (startIndex + 2) > length - 1) {
                                parameter = content;
                            } else {
                                String substring = content.substring(startIndex + 2, endIndex);
                                boolean numeric = StringUtils.isNumeric(substring);
                                if (!numeric) {  //如果获取到的不是一个数字类型，说明可能手动添加了类似于4-1这种格式，不再进行操作
                                    parameter = content;
                                } else {
//                                    substring = "6-"+index+"-"+substring;
//                                    System.out.println(content.indexOf("图"+substring));
                                    parameter = content.replace("图" + substring, "图" + "6-"+index+"-"+substring);
                                }
                            }


                            P p = numberingCreate.createNumberedParagraph(restart, 0, parameter, 100);
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
            } catch (Exception e) {
                e.printStackTrace();
                logger.error("PageContent6生成异常");
            }

        }
    }

    /**
     * 第七章节：正常机组数据
     */
    class PageContent7 {

        public void createPageContent(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle1, List<Map<String, String>> dataList1
                , Map<String, String[]> mapTitle2, List<Map<String, String>> dataList2, Map<String, String[]> mapTitle3, List<Map<String, String>> dataList3) {
            try {
                wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "7 正常机组数据");
                //创建第一个表格
//                createTable(wpMLPackage, mapTitle1, dataList1, 1500);
                Map<String, Object> map1 = new HashMap<>();
                map1.put("firstWidth", 2000);
                map1.put("firstName", "机组编号");
                map1.put("titleMap", mapTitle1);
                map1.put("data", dataList1);
                createTableGeneral(wpMLPackage, map1);
                Docx4jUtil.addBr(wpMLPackage, 10);   //分割表格
                //创建第二个表格
//                createTable(wpMLPackage, mapTitle2, dataList2, 1500);
                Map<String, Object> map2 = new HashMap<>();
                map2.put("firstWidth", 2000);
                map2.put("firstName", "机组编号");
                map2.put("titleMap", mapTitle2);
                map2.put("data", dataList2);
                createTableGeneral(wpMLPackage, map2);
                Docx4jUtil.addBr(wpMLPackage, 2);
                //创建第三个表格
                createTable3(wpMLPackage, mapTitle3, dataList3, 1500);
                //TODO 删
                System.out.println("PageContent7 Success......");
            } catch (Exception e) {
                e.printStackTrace();
                logger.error("PageContent7生成异常");
            }
        }
        private void createTableGeneral(WordprocessingMLPackage wpMLPackage, Map<String, Object> map) {
            //获取参数
            Integer firstWidth = (Integer) map.get("firstWidth");
            firstWidth = firstWidth == null ? -1 : firstWidth;

            String firstName = (String) map.get("firstName");
            firstName = firstName == null ? "机组编号" : firstName;

            Map<String, String[]> titleMap = (Map<String, String[]>) map.get("titleMap");
            if (titleMap == null || titleMap.size() == 0) {
                return;
            }
            List<Map<String, String>> dataList = (List<Map<String, String>>) map.get("data");


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
            String color = "#8064a2";
            //机组编号列固定不变，其他标题动态生成
            TableUtil.addTableTc(titleTr1, firstName, firstWidth, true, "22", "white", color);//该单元格固定
            TableUtil.addTableTc(titleTr2, firstName, firstWidth, true, "22", "white", color);//该单元格固定
            //合并
            TableUtil.mergeCellsVertically(tbl, 0, 0, 1);
            Set<Map.Entry<String, String[]>> entries = titleMap.entrySet();
            //控制第一行合并起始
            int start = 1;
            //记录标题所在位置，后面插入数据时根据该位置将数据插入到对应的标题列下面
            Map<Integer, String> titlePosition = new HashMap<>();
            Integer index = 2;  //除去编号，从第2列开始
            //记录表格列数
            int rowNum = 0;
            for (Map.Entry<String, String[]> entry : entries) {
                if (entry == null) {
                    continue;
                }
                String[] value = entry.getValue();
                if (value == null) {
                    continue;
                }
                rowNum += value.length;
            }
            //单元格宽度
            int width = rowNum == 0 ? -1 : 10000 / rowNum;
//            TableUtil.setTableWidth(tbl, "" + firstWidth + 10000);
            System.out.println("createTable:" + width + "-tableWidth:" + (firstWidth + 10000)); // TODO
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
                    TableUtil.addTableTc(titleTr1, key, width, true, "22", "white", color);
                    //添加第二行标题
                    TableUtil.addTableTc(titleTr2, s, width, true, "22", "white", color);
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
            for (Map<String, String> dataMap : dataList) {
                if (dataMap == null || dataMap.size() == 0) {
                    continue;
                }
                Tr tr = factory.createTr();
                TableUtil.addTableTc(tr, (dataMap.get(firstName) != null && !"".equals(dataMap.get(firstName)) ? dataMap.get(firstName) : "\\"), firstWidth, true, "22", "#000000", null);
                for (int i = 0; i < titlePosition.size(); i++) {//根据位置插入相应的数据
                    TableUtil.addTableTc(tr, (dataMap.get(titlePosition.get(i + 2)) != null && !"".equals(dataMap.get(titlePosition.get(i + 2)))) ? dataMap.get(titlePosition.get(i + 2)) : "\\", width, false, "22", "#000000", null);
                }
                tbl.getContent().add(tr);
            }
        }
        private void createTable3(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList ,int numberIdWidth) {
            Tbl tbl = factory.createTbl();
            //将表格添加到wpMlPackage中
            wpMLPackage.getMainDocumentPart().addObject(tbl);
            //设置样式
            setStyle(tbl);
            //生成表头
            Tr titleTr = factory.createTr();
            tbl.getContent().add(titleTr);
            //机组编号列固定不变，其他标题动态生成
            TableUtil.addTableTc(titleTr, "机组编号", numberIdWidth, true, "22", "white", "#8064a2");//该单元格固定
            Set<Map.Entry<String, String[]>> entries = mapTitle.entrySet();
            //控制第一行合并起始
            int start = 1;
            //记录标题所在位置，后面插入数据时根据该位置将数据插入到对应的标题列下面
            Map<Integer, String> titlePosition = new HashMap<>();
            Integer index = 2;  //除去编号，从第2列开始

            TableUtil.addTableTc(titleTr, "风速(m/s)", 2500, true, "22", "white", "#8064a2");
            titlePosition.put(index++, "风速");

            TableUtil.addTableTc(titleTr, "转速(rpm)", 2500, true, "22", "white", "#8064a2");
            titlePosition.put(index++, "转速");

            TableUtil.addTableTc(titleTr, "功率(kw)", 2500, true, "22", "white", "#8064a2");
            titlePosition.put(index++, "功率");

            //下面插入数据部分
            if (dataList == null || dataList.size() == 0) {
                return;
            }
            for (Map<String, String> map : dataList) {
                if (map == null || map.size() == 0) {
                    continue;
                }
                Tr tr = factory.createTr();
                TableUtil.addTableTc(tr, (map.get("机组编号") != null && !"".equals(map.get("机组编号")) ? map.get("机组编号") : "\\"), numberIdWidth, true, "22", "#000000", null);
                for (int i = 0; i < titlePosition.size(); i++) {//根据位置插入相应的数据
                    TableUtil.addTableTc(tr, (map.get(titlePosition.get(i + 2)) != null && !"".equals(map.get(titlePosition.get(i + 2)))) ? map.get(titlePosition.get(i + 2)) : "\\", 2500, false, "22", "#000000", null);
                }
                tbl.getContent().add(tr);
            }
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
                logger.error("PageContent8生成异常");
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

    /*private void createTable(WordprocessingMLPackage wpMLPackage, Map<String, String[]> mapTitle, List<Map<String, String>> dataList ,int numberIdWidth) {
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
        //设置表格宽度
        TableUtil.setTableWidth(tbl, "" + numberIdWidth + 5000);
        String color = "#8064a2";
        //机组编号列固定不变，其他标题动态生成
        TableUtil.addTableTc(titleTr1, "机组编号", numberIdWidth, true, "22", "white", color);//该单元格固定
        TableUtil.addTableTc(titleTr2, "机组编号", numberIdWidth, true, "22", "white", color);//该单元格固定
        //合并
        TableUtil.mergeCellsVertically(tbl, 0, 0, 1);
        Set<Map.Entry<String, String[]>> entries = mapTitle.entrySet();
        //控制第一行合并起始
        int start = 1;
        //记录标题所在位置，后面插入数据时根据该位置将数据插入到对应的标题列下面
        Map<Integer, String> titlePosition = new HashMap<>();
        Integer index = 2;  //除去编号，从第2列开始
        //记录表格列数
        int rowNum = 0;
        for (Map.Entry<String, String[]> entry : entries) {
            if (entry == null) {
                continue;
            }
            String[] value = entry.getValue();
            if (value == null) {
                continue;
            }
            rowNum += value.length;
        }
        //单元格宽度
        int width = rowNum == 0 ? -1 : 10000 / rowNum;
        TableUtil.setTableWidth(tbl, "" + numberIdWidth + 10000);
        System.out.println("createTable:" + width + "-tableWidth:" + (numberIdWidth + 10000)); // TODO
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
                TableUtil.addTableTc(titleTr1, key, width, true, "22", "white", color);
                //添加第二行标题
                TableUtil.addTableTc(titleTr2, s, width, true, "22", "white", color);
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
            TableUtil.addTableTc(tr, (map.get("机组编号") != null && !"".equals(map.get("机组编号")) ? map.get("机组编号") : "\\"), numberIdWidth, true, "22", "#000000", null);
            for (int i = 0; i < titlePosition.size(); i++) {//根据位置插入相应的数据
                TableUtil.addTableTc(tr, (map.get(titlePosition.get(i + 2)) != null && !"".equals(map.get(titlePosition.get(i + 2)))) ? map.get(titlePosition.get(i + 2)) : "\\", width, false, "22", "#000000", null);
            }
            tbl.getContent().add(tr);
        }
    }*/


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
