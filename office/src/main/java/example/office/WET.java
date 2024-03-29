package example.office;

import example.office.utils.CharsetUtils;
import example.office.word.Word;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.nio.charset.Charset;
import java.util.List;

public class WET {
    private Dict dict;

    public WET(Dict dict) {
        this.dict = dict;
    }

    /**
     * 处理所有文档
     */
    public void process(File doc, File targetDir) throws IOException {
        process("", doc, targetDir);
    }
    public void process(String path, File doc, File targetDir) throws IOException {
        if(doc.isDirectory()) {
            File[] files = doc.listFiles();
            if(files != null) {
                File targetDirChild = new File(targetDir, doc.getName());
                for(File file: files) {
                    process(path + "/" + doc.getName(),file, targetDirChild);
                }
            }
            return;
        }

        System.out.println(path + "/" + doc.getName() + " 开始处理...");
        File targetFile = new File(targetDir, doc.getName());
        targetFile.getParentFile().mkdirs();

        // 读取源文件
        if(doc.getName().endsWith(".doc") || doc.getName().endsWith(".docx")) {
            Word word = new Word(dict);
            word.processDoc(doc, targetFile);
            System.out.println(word);
        } else if(doc.getName().endsWith(".xls") || doc.getName().endsWith(".xlsx")) {
            processXls(doc, targetFile);
        } else if(doc.getName().endsWith(".txt")) {
            processTxt(doc, targetFile);
        } else {
            FileUtils.copyFile(doc, targetFile);
        }

        // 将更改写入目标文件
        System.out.println(path + "/" + doc.getName() + " 完成.");
    }

    /**
     * 替换.txt文件
     */
    private void processTxt(File doc, File targetFile) throws IOException {
        byte[] bytes = FileUtils.readFileToByteArray(doc);
        Charset charset = CharsetUtils.detectCharset(bytes);

        String text = new String(bytes, charset);
        String newText = dict.replaceAll(text);
        FileUtils.writeStringToFile(targetFile, newText, charset);
    }

    /**
     * 替换xls、xlsx文件
     */
    private void processXls(File doc, File targetFile) throws IOException {
        // 读取 Excel 文件
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(doc))) {
            for(Sheet datatypeSheet: workbook) {
                for (Row currentRow : datatypeSheet) {
                    for (Cell cell : currentRow) {
                        String cellValue = TokenConf.cellValue(cell);
                        if (cellValue == null) {
                            continue;
                        }

                        String newCellValue = dict.replaceAll(cellValue);
                        if (!newCellValue.equals(cellValue)) {
                            cell.setCellValue(newCellValue);
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(targetFile)) {
                workbook.write(fos);
            }
        }
    }
}
