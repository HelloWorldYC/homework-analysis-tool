package myc.Entity;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 报告表数据，一个实例存储一张表的信息
 */
@Data
public class HomeworkReport {
    // 班级名
    private String className;
    // 日期
    private String date;
    // 班级人数
    private int studentNum;
    // 平均分
    private double aveScore;
    // 各组成分平均分
    private double[] aveComponentScore;
    // 作业报告表
    private Sheet sheet;
    // 表头
    private String[] header;
    // 表总行数
    private int totalRows;
    // 表总列数
    private int totalCols;
    // 学生列表
    private List<Student> students;
    // 已完成名单
    private List<List<Student>> finishedList;
    // 未完成名单
    private List<List<Student>> unfinishedList;

    /**
     * 有参构造，导入表
     * @param path
     */
    public HomeworkReport(String path){
        try (Workbook workbook = new XSSFWorkbook(path)) {
            Sheet sheet = workbook.getSheetAt(0);
            this.sheet = sheet;
            importExcelInformation();
            importStudentScore();
        } catch (Exception e) {
            System.err.println("获取表" + path + "不成功");
            e.printStackTrace();
        }

    }

    /**
     * 获取报告表的表头信息
     */
    public void importExcelInformation() {
        this.className = sheet.getRow(1).getCell(0).getStringCellValue();
        this.date = sheet.getRow(2).getCell(5).getStringCellValue();
        this.studentNum = countStudentNum();
        totalCols = sheet.getRow(3).getPhysicalNumberOfCells();
        header = new String[totalCols];
        for(int i = 0; i < totalCols; i++) {
            header[i] = sheet.getRow(3).getCell(i).getStringCellValue();
        }
    }

    /**
     * 获取班级学生人数
     */
    public int countStudentNum() {
        int rows = sheet.getPhysicalNumberOfRows();
        int num = rows;
        totalRows = rows;
        for(int i = 0; i < rows; i++) {
            Cell cell = sheet.getRow(i).getCell(0);
            
            if(cell.getCellType() != CellType.NUMERIC) {
                num--;
            }
        }
        // System.out.println("班级人数为：" + num);
        return num;
    }

    /**
     * 导入各学生成绩
     */
    public void importStudentScore() {
        students = new ArrayList<>();
        // 学生成绩组成成分个数
        int componentNum = totalCols - 5;
        finishedList = new ArrayList<>();
        unfinishedList = new ArrayList<>();
        for(int j = 0; j < componentNum; j++) {
            List<Student> finList = new ArrayList<>();
            List<Student> unfinList = new ArrayList<>();
            finishedList.add(finList);
            unfinishedList.add(unfinList);
        }

        DataFormatter dataFormatter = new DataFormatter();
        for(int i = 0; i < studentNum; i++) {
            Student student = new Student();
            Row row = sheet.getRow(i + 4);
            // 设置某学生姓名
            student.setName(row.getCell(1).getStringCellValue());
            // 设置某学生学号
            if(row.getCell(2).getCellType().equals(CellType.NUMERIC)) {
                student.setSchoolId(dataFormatter.formatCellValue(row.getCell(2)));
            }
            // 设置某学生班级
            student.setClassName(className);
            // 设置某学生完成情况
            String[] s = row.getCell(3).getStringCellValue().split("/");
            double[] finishRate = new double[2];
            finishRate[0] = Double.parseDouble(s[0]);
            finishRate[1] = Double.parseDouble(s[1]);
            student.setFinishRate(finishRate[0] / finishRate[1]);
            // 设置某学生总分
            student.setTotalScore(row.getCell(4).getNumericCellValue());

            // 设置某学生各组成分
            double[] scores = new double[componentNum];
            for (int j = 0; j < componentNum; j++) {
                if(row.getCell(j + 5).getCellType().equals(CellType.NUMERIC)) {
                    scores[j] = row.getCell(j + 5).getNumericCellValue();
                    finishedList.get(j).add(student);
                } else {
                    unfinishedList.get(j).add(student);
                }
            }
            student.setScores(scores);

            students.add(student);
        }

        // 设置班级总平均分
        aveScore = sheet.getRow(totalRows - 2).getCell(4).getNumericCellValue();

        // 设置班级分数组成部分平均分
        aveComponentScore = new double[componentNum];
        for (int j = 0; j < componentNum; j++) {
            aveComponentScore[j] = sheet.getRow(totalRows - 2).getCell(j + 5).getNumericCellValue();
        }
    }
}
