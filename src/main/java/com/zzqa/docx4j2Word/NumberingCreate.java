package com.zzqa.docx4j2Word;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.P;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;

import javax.xml.bind.JAXBException;
import java.math.BigInteger;
/**
 * ClassName: NumberingCreate
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/2 10:48
 */


/**
 * Creates a WordprocessingML document from scratch,
 * including a numbering definitions part, and use
 * it to demonstrate restart numbering.
 *
 * @author Jason Harrop
 */
public class NumberingCreate {

    private org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();

    private WordprocessingMLPackage wordMLPackage ;

    private NumberingDefinitionsPart ndp;

    {
        try {
            ndp = new NumberingDefinitionsPart();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }
    //使用原序号
    public void unmarshlDefaultNumbering(){
        try {
            ndp.unmarshalDefaultNumbering();
        } catch (JAXBException e) {
            e.printStackTrace();
        }
    }

    public NumberingCreate(WordprocessingMLPackage wordMLPackage) {
        this.wordMLPackage = wordMLPackage;
        try {
            wordMLPackage.getMainDocumentPart().addTargetPart(ndp);
            ndp.setJaxbElement( (org.docx4j.wml.Numbering) XmlUtils.unmarshalString(initialNumbering) );
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (JAXBException e) {
            e.printStackTrace();
        }

    }

    public long restart(long numID, long ilvl, long val){
        return ndp.restart(numID, ilvl, val);
    }

    /**
     *
     * @return
     */
    public P createNumberedParagraph(long numId, long ilvl, String paragraphText ,int firstLineChars) {

        P  p = factory.createP();

        org.docx4j.wml.Text  t = factory.createText();
        t.setValue(paragraphText);

        org.docx4j.wml.R  run = factory.createR();
        run.getContent().add(t);

        p.getContent().add(run);

        org.docx4j.wml.PPr ppr = factory.createPPr();
        //缩进
        PPrBase.Ind ind = ppr.getInd();
        if (ind==null){
            ind = factory.createPPrBaseInd();
        }
        ind.setFirstLineChars(BigInteger.valueOf(firstLineChars));
        ppr.setInd(ind);
        p.setPPr( ppr );

        // Create and add <w:numPr>
        NumPr numPr =  factory.createPPrBaseNumPr();
        ppr.setNumPr(numPr);

        // The <w:ilvl> element
        Ilvl ilvlElement = factory.createPPrBaseNumPrIlvl();
        numPr.setIlvl(ilvlElement);
        ilvlElement.setVal(BigInteger.valueOf(ilvl));

        // The <w:numId> element
        NumId numIdElement = factory.createPPrBaseNumPrNumId();
        numPr.setNumId(numIdElement);
        numIdElement.setVal(BigInteger.valueOf(numId));

        return p;

    }


    private final String initialNumbering = "<w:numbering xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
            + "<w:abstractNum w:abstractNumId=\"0\">"
            + "<w:nsid w:val=\"2DD860C0\"/>"
            + "<w:multiLevelType w:val=\"multilevel\"/>"
            + "<w:tmpl w:val=\"0409001D\"/>"
            + "<w:lvl w:ilvl=\"0\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"decimal\"/>"
            + "<w:lvlText w:val=\"%1)\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"360\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"1\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"lowerLetter\"/>"
            + "<w:lvlText w:val=\"%2)\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"720\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"2\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"lowerRoman\"/>"
            + "<w:lvlText w:val=\"%3)\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"1080\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"3\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"decimal\"/>"
            + "<w:lvlText w:val=\"(%4)\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"1440\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"4\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"lowerLetter\"/>"
            + "<w:lvlText w:val=\"(%5)\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"1800\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"5\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"lowerRoman\"/>"
            + "<w:lvlText w:val=\"(%6)\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"2160\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"6\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"decimal\"/>"
            + "<w:lvlText w:val=\"%7.\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"2520\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"7\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"lowerLetter\"/>"
            + "<w:lvlText w:val=\"%8.\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"2880\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "<w:lvl w:ilvl=\"8\">"
            + "<w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"lowerRoman\"/>"
            + "<w:lvlText w:val=\"%9.\"/>"
            + "<w:lvlJc w:val=\"left\"/>"
            + "<w:pPr>"
            + "<w:ind w:left=\"3240\" w:hanging=\"360\"/>"
            + "</w:pPr>"
            + "</w:lvl>"
            + "</w:abstractNum>"
            + "<w:num w:numId=\"1\">"
            + "<w:abstractNumId w:val=\"0\"/>"
            + "</w:num>"
            + "</w:numbering>";


}
