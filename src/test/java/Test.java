import java.text.SimpleDateFormat;
import java.util.Date;

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
        System.out.println(simpleDateFormat.format(new Date(1601372015128L)));
    }
}
