package com.zzqa.pojo;

import java.util.List;
import java.util.Map;

/**
 * ClassName: FaultUnitInfo
 * Description:
 *
 * @author 张文豪
 * @date 2020/9/3 14:50
 */
public class FaultUnitInfo {
    private String unitName;
    private String content;
    private String conclusion;
    private List<String[]> imageList;

    public String getUnitName() {
        return unitName;
    }

    public void setUnitName(String unitName) {
        this.unitName = unitName;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getConclusion() {
        return conclusion;
    }

    public void setConclusion(String conclusion) {
        this.conclusion = conclusion;
    }

    public List<String[]> getImageList() {
        return imageList;
    }

    public void setImageList(List<String[]> imageList) {
        this.imageList = imageList;
    }
}
