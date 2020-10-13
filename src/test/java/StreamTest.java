import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

/**
 * ClassName: StreamTest
 * Description:
 *
 * @author 张文豪
 * @date 2020/10/13 9:45
 */
public class StreamTest {
    public static void main(String[] args) {
        List<Singer> singerList = new ArrayList<Singer>();
        singerList.add(new Singer("jay", 11, 36));
        singerList.add(new Singer("eason", 8, 31));
        singerList.add(new Singer("JJ", 6, 29));

        List<String> singerNameList = singerList.stream()
                .filter(singer -> singer.getAge() > 30)  //筛选年龄大于30
                .sorted(Comparator.comparing(Singer::getSongNum))  //根据歌曲数量排序
                .map(Singer::getName)  //提取歌手名字
                .collect(Collectors.toList()); //转换为List
    }
}

class Singer {

    private String name;
    private Integer songNum;
    private Integer age;

    public Singer(String name, Integer songNum, Integer age) {
        this.name = name;
        this.songNum = songNum;
        this.age = age;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getSongNum() {
        return songNum;
    }

    public void setSongNum(Integer songNum) {
        this.songNum = songNum;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}

