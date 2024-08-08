package com.lnbiuc;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.io.InputStream;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        final String wordFilePath = "sign.docx";
        final String imageFilePath = "IMG_20210319_001710.jpg";
        final String outputFilePath = "sign_handled.docx";

        InputStream word = Main.class.getClassLoader().getResourceAsStream(wordFilePath);
        InputStream image = Main.class.getClassLoader().getResourceAsStream(imageFilePath);

        WordPictureReplacer.replaceBookmarkWithImage(word, image, outputFilePath, "sign");
    }
}