package cc.whohow.excel;

import com.fasterxml.jackson.annotation.JsonPropertyDescription;

public class DataModel2 {
    @JsonPropertyDescription("题型（必填）")
    private String type;
    @JsonPropertyDescription("题目（必填）")
    private String stem;
    @JsonPropertyDescription("答案（必填）")
    private String answerKeys;
    @JsonPropertyDescription("选项A（必填）")
    private String answerA;
    @JsonPropertyDescription("选项B（必填）")
    private String answerB;
    @JsonPropertyDescription("选项C（必填）")
    private String answerC;
    @JsonPropertyDescription("选项D（必填）")
    private String answerD;
    @JsonPropertyDescription("选项E（必填）")
    private String answerE;
    @JsonPropertyDescription("选项F（必填）")
    private String answerF;
    @JsonPropertyDescription("标记")
    private String flag;

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getStem() {
        return stem;
    }

    public void setStem(String stem) {
        this.stem = stem;
    }

    public String getAnswerKeys() {
        return answerKeys;
    }

    public void setAnswerKeys(String answerKeys) {
        this.answerKeys = answerKeys;
    }

    public String getAnswerA() {
        return answerA;
    }

    public void setAnswerA(String answerA) {
        this.answerA = answerA;
    }

    public String getAnswerB() {
        return answerB;
    }

    public void setAnswerB(String answerB) {
        this.answerB = answerB;
    }

    public String getAnswerC() {
        return answerC;
    }

    public void setAnswerC(String answerC) {
        this.answerC = answerC;
    }

    public String getAnswerD() {
        return answerD;
    }

    public void setAnswerD(String answerD) {
        this.answerD = answerD;
    }

    public String getAnswerE() {
        return answerE;
    }

    public void setAnswerE(String answerE) {
        this.answerE = answerE;
    }

    public String getAnswerF() {
        return answerF;
    }

    public void setAnswerF(String answerF) {
        this.answerF = answerF;
    }

    public String getFlag() {
        return flag;
    }

    public void setFlag(String flag) {
        this.flag = flag;
    }

    @Override
    public String toString() {
        return "DataModel2{" +
                "type='" + type + '\'' +
                ", stem='" + stem + '\'' +
                ", answerKeys='" + answerKeys + '\'' +
                ", answerA='" + answerA + '\'' +
                ", answerB='" + answerB + '\'' +
                ", answerC='" + answerC + '\'' +
                ", answerD='" + answerD + '\'' +
                ", answerE='" + answerE + '\'' +
                ", answerF='" + answerF + '\'' +
                ", flag='" + flag + '\'' +
                '}';
    }
}
