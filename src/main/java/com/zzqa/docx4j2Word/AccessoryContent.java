package com.zzqa.docx4j2Word;

import com.zzqa.utils.Docx4jUtil;
import com.zzqa.utils.TableUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.util.List;
import java.util.Map;

/**
 * ClassName: AccessoryContent
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/2 15:17
 */
public class AccessoryContent {
    private ObjectFactory factory = new ObjectFactory();

    /**
     * 附件表格信息
     *
     * @param wpMLPackage 传入的wpMLPackage对象
     */
    public void createAccessoryContent(WordprocessingMLPackage wpMLPackage) {
        Docx4jUtil.addNextPage(wpMLPackage);
        //第一个附件标题
        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "附一：典型滚动轴承故障原因分析");
        Docx4jUtil.addParagraph(wpMLPackage, "滚动轴承损伤过程是逐步发展的，且一旦发生，会随着旋转运行而扩散，同时振动会明显增加。引起轴承损伤的原因主要有以下几点：");
        //创建第一个表格
        Tbl tbl1 = factory.createTbl();
        //生成表头
        createTalbeTitle(tbl1);
        //生成数据
        createTableData1(wpMLPackage, tbl1);

        Docx4jUtil.addNextPage(wpMLPackage);
        //第二个附件标题
        wpMLPackage.getMainDocumentPart().addStyledParagraphOfText("Heading1", "附二：典型齿轮箱故障原因分析");
        //创建第一个表格
        Tbl tbl2 = factory.createTbl();
        //生成表头
        createTalbeTitle(tbl2);
        //生成数据
        createTableData2(wpMLPackage, tbl2);
        //TODO 删除输出语句
        System.out.println("AccessoryContent Success......");
        //跨行合并
//        mergeCellsVertically(tbl,0, 1, 2);
    }


    /**
     * 生成表格数据
     *
     * @param wpMLPackage
     * @param tbl
     */
    private void createTableData1(WordprocessingMLPackage wpMLPackage, Tbl tbl) {
        Tr dataTr = null;
        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "轴承磨损", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "安装设计不当", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "没有按正确的安装方法安装轴承，把轴承安装在有椭圆度及锥度的箱体孔时，使轴承产生变形；或者把轴承安装到有椭圆形及锥形的轴颈上时，使轴承也产生变形；轴承外圈与箱体孔配合处落入金属小颗粒，或者内圈与轴顼处落入金属小颗粒，使轴承产生变形；安装轴承外圈的箱体孔中心线对安装轴承内圈的轴中心线歪斜，使轴承在运转中产生别劲”、发热或者磨拟过快；采用加热法安转时，轴承加热温度太高超过规定加热温度，使轴承退火，硬度降低等安装不合理都会造成轴承磨损。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "轴承磨损", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "过负荷", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "轴承过负荷将引起轴承的过早疲劳，使轴承过早磨损而达不到使用设计使用寿命。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "轴承磨损", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "过热", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "轴承温度超过400F时将使滚道和滚动体材料退火，从而使硬度降低导致轴承承重降低和早期失效。严重情况下引起变形，另外温升降低和破坏润滑性能。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "轴承磨损", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "污染", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "当工作环境比较恶劣时，轴承在运行过程中混入杂质从而导致滚道和滚动体表面有点痕，造成轴承振动加大和磨损。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "轴承磨损", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "润滑油失效", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "当轴承润滑油不足或者温度过高时润滑油质量下降时，使轴承各部件之间的油膜遭到破坏，从而使轴承各部件之间相互磨损。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "轴承磨损", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "腐蚀", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "轴承接触腐蚀性流体和气体时会引起轴承早期疲劳磨损。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "轴承磨损", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "配合松动", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "轴承各部件之间的配合松动将导致轴承磨损，并且磨损产生的颗粒将使研磨和松动进一步增大，导致轴承磨损加大。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);
        TableUtil.mergeCellsVertically(tbl, 0, 1, 7);
        wpMLPackage.getMainDocumentPart().addObject(tbl);

    }

    /**
     * 生成表格数据
     *
     * @param wpMLPackage
     * @param tbl
     */
    private void createTableData2(WordprocessingMLPackage wpMLPackage, Tbl tbl) {
        Tr dataTr = null;
        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "齿轮啮合不良", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "齿面磨损", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "齿面磨损主要包括由于灰砂、硬屑粒等进入齿面间而引起的磨粒性磨损和因齿面互相摩擦而产生的跑合性磨损。磨损后齿廓失去正确形状使运转中产生冲击和噪声。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "齿轮啮合不良", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "齿面点蚀", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "相互啮合的两轮齿接触时，齿面间的作用力和反作用力使两工作表面上产生接触应力，由于啮合点的位置是变化的，且齿轮做的是周期性的运动，所以接触应力是按脉动循环变化的。齿面长时间在这种交变接触应力作用下，在齿面的刀痕处会出现小的裂纹，随着时间的推移，这种裂纹逐渐在表层横向扩展，裂纹形成环状后，使轮齿的表面产生微小面积的剥落而形成一些疲劳浅坑。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "齿轮啮合不良", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "齿面胶合", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "在高速重载传动中，常因啮合温度升高而引起润滑失效，致使两齿面金属直接接触并相互粘联。当两齿面相对运动时，较软的齿面沿滑动方向被撕裂出现沟纹，这种现象称为胶合。在低速重载传动中，由于齿面间不易形成润滑油膜也可能产生胶合破坏。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "齿轮啮合不良", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "塑性变形", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "在冲击载荷或重载下，齿面易产生局部的塑性变形，从而使渐开线齿廓的曲面发生变形。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        dataTr = factory.createTr();
        TableUtil.addTableTc(dataTr, "齿轮不对中", 1500, true, "22","black", null);
        TableUtil.addTableTc(dataTr, "装配误差", 1500, false, "22","black", null);
        TableUtil.addTableTc(dataTr, "由于装配技术和装配方法等原因通常会引起齿轮“一端接触”和齿轮轴的直线性偏差（不同轴，不对中）及齿轮的不平衡等故障。", 4500, false, "22","black", null);
        tbl.getContent().add(dataTr);

        TableUtil.mergeCellsVertically(tbl, 0, 1, 4);
        wpMLPackage.getMainDocumentPart().addObject(tbl);

    }

    /**
     * 为table生成表头
     *
     * @param tbl
     */
    private void createTalbeTitle(Tbl tbl) {
        //给table添加边框
        TableUtil.addBorders(tbl);
        Tr tr = factory.createTr();
        //单元格居中对齐
        Jc jc = new Jc();
        jc.setVal(JcEnumeration.CENTER);
        TblPr tblPr = tbl.getTblPr();
        tblPr.setJc(jc);
        tbl.setTblPr(tblPr);

        //表格表头
        TableUtil.addTableTc(tr, "故障类型", 1500, true, "22","black", null);
        TableUtil.addTableTc(tr, "原因", 1500, true, "22","black", null);
        TableUtil.addTableTc(tr, "原因", 4500, true, "22","black", null);
        //将tr添加到table中
        tbl.getContent().add(tr);
        TableUtil.mergeCellsHorizontal(tbl, 0, 1, 2);
    }
}
