package myc.Entity;

import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.Pane;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import javafx.stage.WindowEvent;
import lombok.Data;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

@Data
public class UserInteration {
    private Stage stage;
    private String directory;

    public UserInteration(Stage stage) {
        this.stage = stage;
    }

    /**
     * 交互界面
     * 主要是为了获取用户需要分析的文件的路径，以及开始分析
     */
    public void createUI() {
        stage.setTitle("作业汇总工具");
        // 定义各个组件
        Button selectButton = new Button("选择文件夹");
        Button analysisButton = new Button("开始分析");
        Label label1 = new Label("选择作业报告表所在文件夹");
        Label label2 = new Label();

        // 定义各个组件大小和位置
        selectButton.setPrefSize(80, 30);
        selectButton.setLayoutX(90);
        selectButton.setLayoutY(120);
        analysisButton.setPrefSize(80, 30);
        analysisButton.setLayoutX(230);
        analysisButton.setLayoutY(120);
        label1.setWrapText(true);
        label1.setLayoutX(120);
        label1.setLayoutY(30);
        label2.setWrapText(true);
        label2.setLayoutX(120);
        label2.setLayoutY(70);

        // selectButton 的点击事件
        selectButton.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent mouseEvent) {
                DirectoryChooser directoryChooser = new DirectoryChooser();
                directoryChooser.setTitle("选择一个文件夹");
                File selectedDirectory = directoryChooser.showDialog(stage);
                if(selectedDirectory != null) {
                    System.out.println("已选择文件夹：" + selectedDirectory.getAbsolutePath());
                    directory = selectedDirectory.getAbsolutePath();
                    label2.setText("已选择文件夹：\n" + selectedDirectory.getAbsolutePath());
                }
            }
        });

        // analysisButton 的点击事件
        analysisButton.setOnMouseClicked(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent mouseEvent) {
                if(directory == null || directory.length() == 0) {
                    label2.setText("没有文件夹被选择！");
                    return;
                } else {
                    File file = new File(directory);
                    String[] fileList = file.list();
                    List<String> inputPaths = new ArrayList<>();
                    for (int i = 0; i < fileList.length; i++) {
                        if(fileList[i].endsWith(".xlsx") && !"汇总报告表.xlsx".equals(fileList[i])) {
                            inputPaths.add(directory + "\\" + fileList[i]);
                        }
                    }
                    String outputPath = directory + "\\" + "汇总报告表.xlsx";
                    ReportAnalysis reportAnalysis = new ReportAnalysis(outputPath);
                    for(String s : inputPaths) {
                        HomeworkReport homeworkReport = new HomeworkReport(s);
                        reportAnalysis.addHomeworkReport(homeworkReport);
                    }
                    reportAnalysis.exportSummarizationExcel();
                    reportAnalysis.exportTotalStudentExcel();
                    reportAnalysis.exportFinishedExcel();
                    reportAnalysis.exportUnfinishedExcel();
                    label2.setText("分析已经完成！结果汇总表在：\n" + outputPath);
                    System.out.println("分析已经完成！结果汇总表在：" + outputPath);
                }

            }
        });

        Pane pane = new Pane();
        pane.getChildren().addAll(selectButton, analysisButton, label1, label2);
        Scene scene = new Scene(pane, 400, 200);
        stage.setScene(scene);
        stage.setResizable(false);

        stage.show();
        stage.setOnCloseRequest(new EventHandler<WindowEvent>() {
            @Override
            public void handle(WindowEvent windowEvent) {
                System.out.println("程序执行结束!");
            }
        });
    }


}
