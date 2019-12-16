import excel.Parser;

import java.io.*;

public class Application {
    public static void main(String[] args) throws IOException {
        File currentdir = new File(".");
        File[] files = currentdir.listFiles();
        for (File inst: files) {
            String filename = inst.getName().toLowerCase();
            if (inst.isFile() && filename.endsWith(".xls")
                    || filename.endsWith(".xlsx")) {
                String outname = filename.endsWith(".xls")
                        ? filename.replace(".xls", ".csv")
                        : filename.replace(".xlsx", "csv");
                File out = new File(outname);
                try (FileWriter fileWriter = new FileWriter(out)){
                    fileWriter.write(Parser.parse(inst.getName()));
                    fileWriter.flush();
                }
            }
        }

    }
}
