package com.example.helper.compression;

import net.coobird.thumbnailator.Thumbnails;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;

import javax.imageio.stream.ImageOutputStreamImpl;
import java.io.*;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ImageCompressionDocx {

    public static void main(String[] args) {
        // Nén ảnh trong file DOCX
        compressImages("F:\\test\\sample4.docx", "F:\\test\\zxc.docx");

        // Giải nén ảnh trong file DOCX
//        decompressImages("F:\\test\\zxc.docx", "F:\\test\\output_images");
    }

    // Nén ảnh trong file DOCX
    public static void compressImages(String inputFilePath, String outputFilePath) {
        try (ZipArchiveInputStream zis = new ZipArchiveInputStream(new FileInputStream(inputFilePath));
             ZipArchiveOutputStream zos = new ZipArchiveOutputStream(new FileOutputStream(outputFilePath))) {

            ZipArchiveEntry entry;

            while ((entry = zis.getNextZipEntry()) != null) {
                String entryName = entry.getName();

                // Nếu là ảnh, thực hiện nén
                if (entryName.startsWith("word/media/") && StringUtils.endsWithAny(entryName.toLowerCase(), ".jpeg",".jpg",".png")) {
                    String extension = FilenameUtils.getExtension(entryName);
                    ByteArrayOutputStream compressedImage = compressImage(zis.readAllBytes(), extension);
                    ZipArchiveEntry newEntry = new ZipArchiveEntry(entryName);
                    zos.putArchiveEntry(newEntry);
                    zos.write(compressedImage.toByteArray());
                    zos.closeArchiveEntry();
                } else {
                    // Nếu không phải là ảnh, giữ nguyên entry
                    ZipArchiveEntry newEntry = new ZipArchiveEntry(entry);
                    zos.putArchiveEntry(newEntry);
                    zos.write(zis.readAllBytes());
                    zos.closeArchiveEntry();
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Giải nén ảnh trong file DOCX
    public static void decompressImages(String inputFilePath, String outputDirectory) {
        try (ZipFile zipFile = new ZipFile(inputFilePath)) {
            Enumeration<? extends ZipEntry> entries = zipFile.entries();

            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                String entryName = entry.getName();

                // Nếu là ảnh, thực hiện giải nén
                if (entryName.startsWith("word/media/") && StringUtils.endsWithAny(entryName.toLowerCase(), ".jpeg",".jpg",".png")) {
                    InputStream compressedImage = zipFile.getInputStream(entry);
                    byte[] decompressedImage = decompressImage(compressedImage);

                    // Lưu ảnh đã giải nén vào thư mục đích
                    String outputFilePath = outputDirectory + File.separator + entryName;
                    try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                        fos.write(decompressedImage);
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Nén ảnh
    private static ByteArrayOutputStream compressImage(byte[] imageData, String formatExtension) throws IOException {
        // Đọc ảnh từ byte array
        InputStream inputStream = new ByteArrayInputStream(imageData);

        // Tạo ByteArrayOutputStream để lưu ảnh đã nén
        ByteArrayOutputStream compressedImageStream = new ByteArrayOutputStream();

        // Nén ảnh với Thumbnails
        Thumbnails.of(inputStream)
                .size(640, 480)
                .outputQuality(0.7) // Chất lượng nén (0.7 = 70%)
                .toOutputStream(compressedImageStream);

        return compressedImageStream;
    }

    // Giải nén ảnh
    private static byte[] decompressImage(InputStream compressedImageData) throws IOException {
        // Tạo ByteArrayOutputStream để lưu ảnh đã giải nén
        ByteArrayOutputStream decompressedImageStream = new ByteArrayOutputStream();

        // Giải nén ảnh với Thumbnails
        Thumbnails.of(compressedImageData).scale(1.0).toOutputStream(decompressedImageStream);

        return decompressedImageStream.toByteArray();
        
        // Thực hiện thuật toán giải nén ảnh tại đây
        // Trong ví dụ này, đơn giản chỉ giữ nguyên ảnh không thay đổi
//        return compressedImageData.readAllBytes();
    }


    private static class MyImageOutputStream extends ImageOutputStreamImpl {
        private ByteArrayOutputStream outputStream;

        MyImageOutputStream(ByteArrayOutputStream outputStream) {
            this.outputStream = outputStream;
        }

        @Override
        public void write(int b) throws IOException {
            outputStream.write(b);
        }

        @Override
        public void write(byte[] b, int off, int len) throws IOException {

        }

        @Override
        public int read() throws IOException {
            return 0;
        }

        @Override
        public int read(byte[] b, int off, int len) throws IOException {
            return 0;
        }

        // Các phương thức còn lại của ImageOutputStreamImpl
    }
}
