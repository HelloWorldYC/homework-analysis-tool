package myc.Entity;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

@Data
public class ReportAnalysis {

    private List<HomeworkReport> homeworkReportList;
    private String outputPath;
    private Workbook workbook;
    private ExcelStyleUtil excelStyleUtil;

    public ReportAnalysis(String path) {
        this.homeworkReportList = new ArrayList<>();
        this.outputPath = path;
        this.workbook = new XSSFWorkbook();
        this.excelStyleUtil = new ExcelStyleUtil(workbook);
    }

    public void addHomeworkReport(HomeworkReport homeworkReport) {
        homeworkReportList.add(homeworkReport);
    }

    /**
     * 导出汇总表
     */
    public void exportSummarizationExcel() {
        Sheet summarizationExcel = workbook.createSheet("汇总");
        HomeworkReport homeworkReport = homeworkReportList.get(0);
        int componentNum = homeworkReport.getAveComponentScore().length;
        int totalCol = 3 + 3*componentNum;

        // 第一行
        Row row = summarizationExcel.createRow(0);
        row.createCell(0).setCellValue("成绩统计表-" + homeworkReport.getDate());
        row.getCell(0).setCellStyle(excelStyleUtil.getFirstHeaderStyle());
        row.setHeightInPoints(30);

        // 第二行
        row = summarizationExcel.createRow(1);
        row.setRowStyle(excelStyleUtil.getSecondHeaderStyle());
        Cell cell = row.createCell(0);
        cell.setCellValue("班级");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(1);
        cell.setCellValue("班级人数");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(2);
        cell.setCellValue("完成人数");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(2 + componentNum);
        cell.setCellValue("完成率");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(2 + 2*componentNum);
        cell.setCellValue("总平均分");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(3 + 2*componentNum);
        cell.setCellValue("各组成部分平均分");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());


        // 第三行
        row = summarizationExcel.createRow(2);
        for (int j = 0; j < componentNum; j++) {
            String s = homeworkReport.getHeader()[5 + j];
            cell = row.createCell(2 + j);
            cell.setCellValue(s);
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
            cell = row.createCell(2 + componentNum + j);
            cell.setCellValue(s);
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
            cell = row.createCell(3 + 2*componentNum + j);
            cell.setCellValue(s);
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        }

        // 合并单元格
        summarizationExcel.addMergedRegion(new CellRangeAddress(0, 0, 0, totalCol - 1));
        summarizationExcel.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
        summarizationExcel.addMergedRegion(new CellRangeAddress(1, 2, 1, 1));
        if(componentNum > 1) summarizationExcel.addMergedRegion(new CellRangeAddress(1, 1, 2, 2 + componentNum - 1));
        if(componentNum > 1) summarizationExcel.addMergedRegion(new CellRangeAddress(1, 1, 2 + componentNum, 2 + 2*componentNum - 1));
        summarizationExcel.addMergedRegion(new CellRangeAddress(1, 2, 2 + 2*componentNum, 2 + 2*componentNum));
        if(componentNum > 1) summarizationExcel.addMergedRegion(new CellRangeAddress(1, 1, 3 + 2*componentNum, totalCol - 1));

        // 设置各列宽
        for (int j = 0; j < totalCol; j++) {
            summarizationExcel.setColumnWidth(j, 20 * 256);
        }

        // 统计各班数据
        int classNum = homeworkReportList.size();
        int totalStudentNum = 0;
        int[] totalFinishNum = new int[componentNum];
        double totalScore = 0.0;
        double[] totalComponentScore = new double[componentNum];
        for(int i = 0; i < classNum; i++) {
            homeworkReport = homeworkReportList.get(i);
            row = summarizationExcel.createRow(i + 3);
            // 班级名
            cell = row.createCell(0);
            cell.setCellValue(homeworkReport.getClassName());
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            // 班级人数
            cell = row.createCell(1);
            cell.setCellValue(String.valueOf(homeworkReport.getStudentNum()));
            totalStudentNum += homeworkReport.getStudentNum();
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            // 班级总平均分
            cell = row.createCell(2 + 2*componentNum);
            cell.setCellValue(homeworkReport.getAveScore());
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            totalScore += (homeworkReport.getAveScore()*homeworkReport.getStudentNum());
            for (int j = 0; j < componentNum; j++) {
                // 各分数组成部分完成人数
                cell = row.createCell(2 + j);
                int finishNum = homeworkReport.getFinishedList().get(j).size();
                cell.setCellValue(String.valueOf(finishNum));
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
                totalFinishNum[j] += finishNum;
                // 各分数组成部分完成率
                cell = row.createCell(2 + componentNum + j);
                double finishiRate = (double)finishNum / homeworkReport.getStudentNum();
                cell.setCellValue(finishiRate);
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
                // 各分数组成部分平均分
                cell = row.createCell(3 + 2*componentNum + j);
                cell.setCellValue(homeworkReport.getAveComponentScore()[j]);
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
                totalComponentScore[j] += (homeworkReport.getAveComponentScore()[j] * homeworkReport.getStudentNum());
            }
        }

        // 汇总各班数据
        row = summarizationExcel.createRow(3 + classNum);
        cell = row.createCell(0);
        cell.setCellValue("汇总");
        cell.setCellStyle(excelStyleUtil.getSummarizationRowStyle());
        cell = row.createCell(1);
        cell.setCellValue(String.valueOf(totalStudentNum));
        cell.setCellStyle(excelStyleUtil.getSummarizationRowStyle());
        // 总平均分
        cell = row.createCell(2 + 2*componentNum);
        cell.setCellValue(totalScore / totalStudentNum);
        cell.setCellStyle(excelStyleUtil.getSummarizationRowStyle());
        for (int j = 0; j < componentNum; j++) {
            // 各分数组成部分完成人数
            cell = row.createCell(2 + j);
            cell.setCellValue(String.valueOf(totalFinishNum[j]));
            cell.setCellStyle(excelStyleUtil.getSummarizationRowStyle());
            // 各分数组成部分完成率
            cell = row.createCell(2 + componentNum + j);
            double finishiRate = (double)totalFinishNum[j] / totalStudentNum;
            cell.setCellValue(finishiRate);
            cell.setCellStyle(excelStyleUtil.getSummarizationRowStyle());
            // 各分数组成部分平均分
            cell = row.createCell(3 + 2*componentNum + j);
            double totalAveComponentScore = totalComponentScore[j] / totalStudentNum;
            cell.setCellValue(totalAveComponentScore);
            cell.setCellStyle(excelStyleUtil.getSummarizationRowStyle());
        }

        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
        } catch (Exception e) {
            System.err.println("汇总表导出出错！");
        }
    }

    /**
     * 导出学生名单总表
     */
    public void exportTotalStudentExcel() {
        List<Student> studentList = new ArrayList<>();
        int classNum = homeworkReportList.size();
        for (int i = 0; i < classNum; i++) {
            HomeworkReport homeworkReport = homeworkReportList.get(i);
            List<Student> students = homeworkReport.getStudents();
            studentList.addAll(students);
        }
        // 将学生总名单按总分排序
        Collections.sort(studentList, new Comparator<Student>() {
            @Override
            public int compare(Student o1, Student o2) {
                if(o1.getTotalScore() > o2.getTotalScore()) return -1;
                else if (o1.getTotalScore() == o2.getTotalScore()) return 0;
                return 1;
            }
        });

        // 建表
        Sheet totalStudentExcel = workbook.createSheet("学生总表");
        int componentNum = homeworkReportList.get(0).getAveComponentScore().length;

        // 设置各列宽
        for (int j = 0; j < 5 + componentNum; j++) {
            totalStudentExcel.setColumnWidth(j, 20 * 256);
        }

        // 第一行
        Row row = totalStudentExcel.createRow(0);
        row.createCell(0).setCellValue("学生成绩总表-" + homeworkReportList.get(0).getDate());
        row.getCell(0).setCellStyle(excelStyleUtil.getFirstHeaderStyle());
        row.setHeightInPoints(30);
        totalStudentExcel.addMergedRegion(new CellRangeAddress(0, 0, 0, 4 + componentNum));

        // 第二行
        row = totalStudentExcel.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("序号");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(1);
        cell.setCellValue("班级");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(2);
        cell.setCellValue("姓名");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(3);
        cell.setCellValue("学号");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        cell = row.createCell(4);
        cell.setCellValue("总分");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        for (int j = 0; j < componentNum; j++) {
            cell = row.createCell(5 + j);
            String s = homeworkReportList.get(0).getHeader()[5 + j];
            cell.setCellValue(s);
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        }

        // 导入学生数据
        for (int i = 0; i < studentList.size(); i++) {
            Student student = studentList.get(i);
            row = totalStudentExcel.createRow(i + 2);
            cell = row.createCell(0);
            cell.setCellValue(String.valueOf(i + 1));
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            cell = row.createCell(1);
            cell.setCellValue(student.getClassName());
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            cell = row.createCell(2);
            cell.setCellValue(student.getName());
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            cell = row.createCell(3);
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            if(student.getSchoolId() != null && student.getSchoolId().length() != 0) {
                cell.setCellValue(student.getSchoolId());
            }
            cell = row.createCell(4);
            cell.setCellValue(student.getTotalScore());
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            for (int j = 0; j < componentNum; j++) {
                cell = row.createCell(5 + j);
                cell.setCellValue(student.getScores()[j]);
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            }
        }

        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
        } catch (Exception e) {
            System.err.println("学生总表导出出错！");
        }
    }

    /**
     * 导出已完成学生表
     */
    public void exportFinishedExcel(){
        Sheet finishedExcel = workbook.createSheet("已完成");
        int componentNum = homeworkReportList.get(0).getAveComponentScore().length;

        // 第一行
        Row row = finishedExcel.createRow(0);
        row.createCell(0).setCellValue("已完成的学生名单-" + homeworkReportList.get(0).getDate());
        row.getCell(0).setCellStyle(excelStyleUtil.getFirstHeaderStyle());
        row.setHeightInPoints(30);
        finishedExcel.addMergedRegion(new CellRangeAddress(0, 0, 0, 2 * componentNum));

        // 第二行
        row = finishedExcel.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("班级");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        for (int j = 0; j < 2 * componentNum; j += 2) {
            cell = row.createCell(1 + j);
            cell.setCellValue(homeworkReportList.get(0).getHeader()[5 + j / 2] + "已完成人数");
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
            cell = row.createCell(2 + j);
            cell.setCellValue(homeworkReportList.get(0).getHeader()[5 + j / 2] + "已完成学生名单");
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        }

        // 设置各列宽
        int firstColumnWidth = 20 * 256;
        int finishedNumColumnWidth = 40 * 256;
        int finishedListColumnWidth = 50 * 256;
        finishedExcel.setColumnWidth(0, firstColumnWidth);
        for (int j = 1; j < 1 + 2 * componentNum; j += 2) {
            finishedExcel.setColumnWidth(j, finishedNumColumnWidth);
            finishedExcel.setColumnWidth(j + 1, finishedListColumnWidth);
        }

        // 导入每个班级完成情况
        for (int i = 0; i < homeworkReportList.size(); i++) {
            row = finishedExcel.createRow(2 + i);
            HomeworkReport homeworkReport = homeworkReportList.get(i);
            cell = row.createCell(0);
            cell.setCellValue(homeworkReport.getClassName());
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            List<List<Student>> finishedList = homeworkReport.getFinishedList();
            for (int j = 0; j < 2 * componentNum; j += 2) {
                List<Student> studentList = finishedList.get(j / 2);
                cell = row.createCell(1 + j);
                cell.setCellValue(String.valueOf(studentList.size()));
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
                cell = row.createCell(2 + j);
                StringBuilder sb = new StringBuilder();
                for(Student student : studentList) {
                    sb.append(student.getName());
                    sb.append(" ");
                }
                cell.setCellValue(sb.toString());
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            }
        }

        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
        } catch (Exception e) {
            System.err.println("学生已完成表导出出错！");
        }
    }

    /**
     * 导出未完成学生表
     */
    public void exportUnfinishedExcel(){
        Sheet unfinishedExcel = workbook.createSheet("未完成");
        int componentNum = homeworkReportList.get(0).getAveComponentScore().length;

        // 第一行
        Row row = unfinishedExcel.createRow(0);
        row.createCell(0).setCellValue("未完成的学生名单-" + homeworkReportList.get(0).getDate());
        row.getCell(0).setCellStyle(excelStyleUtil.getFirstHeaderStyle());
        row.setHeightInPoints(30);
        unfinishedExcel.addMergedRegion(new CellRangeAddress(0, 0, 0, 2 * componentNum));

        // 第二行
        row = unfinishedExcel.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("班级");
        cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        for (int j = 0; j < 2 * componentNum; j += 2) {
            cell = row.createCell(1 + j);
            cell.setCellValue(homeworkReportList.get(0).getHeader()[5 + j / 2] + "未完成人数");
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
            cell = row.createCell(2 + j);
            cell.setCellValue(homeworkReportList.get(0).getHeader()[5 + j / 2] + "未完成学生名单");
            cell.setCellStyle(excelStyleUtil.getSecondHeaderStyle());
        }

        // 设置各列宽
        int firstColumnWidth = 20 * 256;
        int finishedNumColumnWidth = 40 * 256;
        int finishedListColumnWidth = 50 * 256;
        unfinishedExcel.setColumnWidth(0, firstColumnWidth);
        for (int j = 1; j < 1 + 2 * componentNum; j += 2) {
            unfinishedExcel.setColumnWidth(j, finishedNumColumnWidth);
            unfinishedExcel.setColumnWidth(j + 1, finishedListColumnWidth);
        }

        // 导入每个班级完成情况
        for (int i = 0; i < homeworkReportList.size(); i++) {
            row = unfinishedExcel.createRow(2 + i);
            HomeworkReport homeworkReport = homeworkReportList.get(i);
            cell = row.createCell(0);
            cell.setCellValue(homeworkReport.getClassName());
            cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            List<List<Student>> unfinishedList = homeworkReport.getUnfinishedList();
            for (int j = 0; j < 2 * componentNum; j += 2) {
                List<Student> studentList = unfinishedList.get(j / 2);
                cell = row.createCell(1 + j);
                cell.setCellValue(String.valueOf(studentList.size()));
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
                cell = row.createCell(2 + j);
                StringBuilder sb = new StringBuilder();
                for(Student student : studentList) {
                    sb.append(student.getName());
                    sb.append(" ");
                }
                cell.setCellValue(sb.toString());
                cell.setCellStyle(excelStyleUtil.getDataCellStyle());
            }
        }

        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            workbook.write(fos);
        } catch (Exception e) {
            System.err.println("学生未完成表导出出错！");
        }
    }
}
