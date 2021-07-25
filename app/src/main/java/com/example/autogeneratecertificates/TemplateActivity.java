package com.example.autogeneratecertificates;

import android.content.Context;
import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.widget.TextView;

import androidx.appcompat.app.AppCompatActivity;

import com.tom_roush.pdfbox.pdmodel.PDDocument;
import com.tom_roush.pdfbox.pdmodel.PDPage;
import com.tom_roush.pdfbox.pdmodel.PDPageContentStream;
import com.tom_roush.pdfbox.pdmodel.PDPageTree;
import com.tom_roush.pdfbox.pdmodel.common.PDStream;
import com.tom_roush.pdfbox.pdmodel.font.PDType1Font;
import com.tom_roush.pdfbox.util.Matrix;
import com.tom_roush.pdfbox.util.PDFBoxResourceLoader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;


public class TemplateActivity extends AppCompatActivity {

    TextView test;
    String[] a=new String[30];
    String[] b=new String[30];
    String[] c=new String[30];
    String[] d=new String[30];
    String[] e=new String[30];
    String[] name=new String[30];
    String[] course=new String[30];
    String[] college=new String[30];
    String[] position=new String[30];
    String[] society=new String[30];

    int row_num;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_template);
        test = (TextView) findViewById(R.id.test);
        try {
            row_num = Integer.parseInt(getIntent().getStringExtra("entries"));
        }catch (NullPointerException e){
            e.printStackTrace();
        }
        readExcelData();
        test.setText("Certificate Generation Completed!");
    }

    public void readExcelData() {
        new Thread(() -> {
            try {
                String path_xlsx = getIntent().getStringExtra("path");
                InputStream inputfile = getContentResolver().openInputStream(Uri.parse(path_xlsx));
                XSSFWorkbook workbook = new XSSFWorkbook(inputfile);
                XSSFSheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();
                int count = 0;
                String temp;
                while (rowIterator.hasNext() && count < row_num+1) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cx = row.cellIterator();
                    temp=cx.next().getStringCellValue();
                    a[count]=temp;

                    temp=cx.next().getStringCellValue();
                    b[count]=temp;

                    temp=cx.next().getStringCellValue();
                    c[count]=temp;

                    temp=cx.next().getStringCellValue();
                    d[count]=temp;

                    temp=cx.next().getStringCellValue();
                    e[count]=temp;

                    count++;
                }
                inputfile.close();
                matchColumn();
                searchSlide();
            } catch (FileNotFoundException e) {
                test.setText("FilenotFound");
                e.printStackTrace();
            } catch (IOException e) {
                test.setText("IOException");
                e.printStackTrace();
            }
        }).start();
    }

    public void matchColumn() {
        int i;
        switch (a[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = a[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=a[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=a[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=a[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=a[i];
                }
                break;
        }
        switch (b[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = b[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=b[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=b[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=b[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=b[i];
                }
                break;
        }
        switch (c[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = c[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=c[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=c[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=c[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=c[i];
                }
                break;
        }
        switch (d[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = d[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=d[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=d[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=d[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=d[i];
                }
                break;
        }
        switch (e[0].toLowerCase()) {
            case "name":
                for (i = 1; i <= row_num; i++) {
                    name[i-1] = e[i];
                }
                break;
            case "college":
                for (i=1; i<=row_num; i++){
                    college[i-1]=e[i];
                }
                break;
            case "course":
                for (i=1; i<=row_num; i++){
                    course[i-1]=e[i];
                }
                break;
            case "society":
                for (i=1; i<=row_num; i++){
                    society[i-1]=e[i];
                }
                break;
            case "position":
                for (i=1; i<=row_num; i++){
                    position[i-1]=e[i];
                }
                break;
        }
    }

    private void searchSlide() {
        try {
            PDFBoxResourceLoader.init(getApplicationContext());
            String encoding = "ISO-8859-1";
            AssetManager assManager = getAssets();

            for (int i=0; i<row_num; i++){
                InputStream is = assManager.open("template1.pdf");
                OutputStream newPDFfile = createFile(name[i]);
                copy(is, newPDFfile);
                File PDFfile = new File(getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)+"/AutoCertiGen/" +name[i]+".pdf");
                PDDocument pdf = PDDocument.load(PDFfile);

                PDPage page = pdf.getPage(0);
                PDPageContentStream contentStream = new PDPageContentStream(pdf, page,PDPageContentStream.AppendMode.APPEND,true,true);
                contentStream.beginText();
                //contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_BOLD, 150);
                contentStream.newLineAtOffset(1400, 1100);
                contentStream.showText(name[i]);
                contentStream.endText();
                contentStream.beginText();
                //contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 120);
                contentStream.newLineAtOffset(1150, 1400);
                contentStream.showText("from "+college[i]+" has secured");
                contentStream.endText();
                contentStream.beginText();
                //contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 120);
                contentStream.newLineAtOffset(1150, 1520);
                contentStream.showText(position[i]+" position in "+society[i]+".");
                contentStream.endText();
                contentStream.close();
                pdf.save(PDFfile);
                pdf.close();
                newPDFfile.close();
                is.close();
            }

            Log.d("TemplateActivity", "Completed");
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    private OutputStream createFile(String name) throws IOException {
        File f = new File(getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)+"/AutoCertiGen/" +name+".pdf");
        if (f.exists()){
            f.delete();
        }
        if (!f.getParentFile().exists()){
            f.getParentFile().mkdirs();
        }
        f.createNewFile();
        OutputStream newFile = new FileOutputStream(f);
        return newFile;
    }

    public static void copy(InputStream fis, OutputStream fos)
            throws IOException
    {
        byte[] buffer = new byte[1024];
        int read;
        while((read = fis.read(buffer)) != -1){
            fos.write(buffer, 0, read);
        }
    }
}