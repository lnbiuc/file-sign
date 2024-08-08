package com.lnbiuc;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;

import java.io.*;
import java.util.List;

/**
 * @Description: TODO
 * @Date: 2024/8/8 16:58
 * @Created: by ZZSLL
 */

public class WordPictureReplacer {
    /**
     * 在 Word 文档中的书签位置插入图片，并保存修改后的文档。
     *
     * @param wordFile     原始 Word 文档路径
     * @param imageFile    图片文件路径
     * @param outputFilePath   修改后的 Word 文档保存路径
     * @return 修改后的 Word 文件
     * @throws IOException      IO 异常
     * @throws InvalidFormatException 格式异常
     */
    public static File replaceBookmarkWithImage(InputStream wordFile, InputStream imageFile, String outputFilePath, String bookmarkName) throws IOException, InvalidFormatException {
        byte[] imageBytes = readImageFile(imageFile);

        try (XWPFDocument doc = new XWPFDocument(wordFile)) {
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                List<CTBookmark> bookmarkStartList = paragraph.getCTP().getBookmarkStartList();
                for (CTBookmark bookmark : bookmarkStartList) {
                    if (bookmarkName.equals(bookmark.getName())) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        XWPFRun run;
                        if (!runs.isEmpty()) {
                            run = runs.get(0);
                        } else {
                            run = paragraph.createRun();
                        }
                        run.addPicture(new ByteArrayInputStream(imageBytes), XWPFDocument.PICTURE_TYPE_PNG, "imageName.png", Units.toEMU(80), Units.toEMU(80));
                    }
                }
            }

            File outputFile = new File(outputFilePath);
            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                doc.write(out);
            }

            return outputFile;
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    private static byte[] readImageFile(InputStream file) throws IOException {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = file.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }
            return outputStream.toByteArray();
        }
    }
}
