import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;

import java.awt.*;
import java.io.IOException;
import java.net.URI;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * ClassName: Test
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/29 17:51
 */
public class Test {
    @org.junit.Test
    public void test(){
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat();
        simpleDateFormat.applyPattern("yyyy-MM-dd HH:mm:ss");
        System.out.println(simpleDateFormat.format(new Date(1602480817606l)));
    }
    @org.junit.Test
    public void test1(){
        String content = "1、发电机自由端轴承振动有效值较大时已达到2.8g（图3）；";
        System.out.println(content.indexOf("（图"));
        System.out.println(content.lastIndexOf("（图"));
        System.out.println(content.lastIndexOf("）"));
        System.out.println(content.substring(content.lastIndexOf("（图")+1,content.lastIndexOf("）")));
        System.out.println(NumberUtils.isDigits(content.substring(content.lastIndexOf("（图")+1,content.lastIndexOf("）"))));
        System.out.println(StringUtils.isNumeric(content.substring(content.lastIndexOf("（图") + 2, content.lastIndexOf("）"))));

        int length = content.length();
        int startIndex = content.lastIndexOf("（图")!=-1?content.lastIndexOf("（图"):content.lastIndexOf("(图");
        int endIndex = content.lastIndexOf("）")!=-1?content.lastIndexOf("）"):content.lastIndexOf(")");
        if (length<(startIndex+2)){
            System.out.println("越界");
        }
        System.out.println(startIndex+"-"+endIndex);
        System.out.println(content.substring(startIndex + 2, endIndex));

    }

    @org.junit.Test
    public void test2() throws IOException, InterruptedException {
        Desktop desktop = Desktop.getDesktop();
        desktop.browse(URI.create("https://www.baidu.com"));
    }
    @org.junit.Test
    public void test3(){
        System.out.println(5/3);
        System.out.println((Integer)(5/3));
    }
}
