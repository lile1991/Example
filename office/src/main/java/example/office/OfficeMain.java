package example.office;

import java.io.File;
import java.io.IOException;

public class OfficeMain {
    public static void main(String[] args) throws IOException {
        // Scanner scanner = new Scanner(System.in);
        // System.out.println("Input dictionary file");
        // String dictFileStr = scanner.nextLine();
        // System.out.println("Enter document path");
        // String docDirStr = scanner.nextLine();

        String docDirStr = "D:\\Workspace\\Example\\office\\src\\test\\resources\\段落测试";
        String docLibDirStr = "D:\\Workspace\\Example\\office\\src\\test\\resources\\素材";
        String dictFileStr = "D:\\Workspace\\Example\\office\\src\\test\\resources\\conf.xlsx";

        File dictFile = new File(dictFileStr);
        File docDir = new File(docDirStr);


        try (Dict dict = new Dict(dictFile)) {
            File docDirTarget = new File(docDir.getParent(), "Target");
            WET WET = new WET(dict);

            WET.process(docDir, docDirTarget);
        }
    }


}
