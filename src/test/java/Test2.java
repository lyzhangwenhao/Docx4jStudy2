import org.junit.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * ClassName: Test2
 * Description:
 *
 * @author 张文豪
 * @date 2020/10/13 9:11
 */
public class Test2 {
    @Test
    public void test1() throws IOException {
        Path path = Paths.get("F:\\AutoExport\\docx4j2\\上海东滩风电场2020年10月12日震动分析报告.docx");
        System.out.println("FileName:"+path.getFileName());
        System.out.println("FileSystem:"+path.getFileSystem());
        System.out.println("Root:"+path.getRoot());
        System.out.println("AbsolutePath:"+path.toAbsolutePath());
        byte[] bytes = Files.readAllBytes(path);
        System.out.println(bytes.length+"---"+bytes.toString());
    }
}
