package myc;


import javafx.application.Application;
import javafx.stage.Stage;
import myc.Entity.UserInteration;


public class HomeworkAnalysis extends Application {

    public static void main(String[] args){
        launch(args);
    }

    @Override
    public void start(Stage stage) throws Exception {
        System.out.println("程序正在执行...");
        UserInteration userInteration = new UserInteration(stage);
        userInteration.createUI();
    }

}
