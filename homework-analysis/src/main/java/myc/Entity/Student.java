package myc.Entity;

import lombok.Data;

@Data
public class Student {
    // 姓名
    private String name;
    // 学号
    private String schoolId;
    // 班级
    private String className;
    // 总分
    private double totalScore;
    // 各组成分数
    private double[] scores;
    // 完成情况
    private double finishRate;
}
