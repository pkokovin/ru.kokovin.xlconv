import excel.Parser;

import java.io.*;

public class Application {
    public static void main(String[] args) throws IOException {
        File out = new File("toparce.csv");
        FileWriter fileWriter = new FileWriter(out);
        fileWriter.write(Parser.parse("toparce.xls"));
        fileWriter.flush();
        fileWriter.close();
    }
}
