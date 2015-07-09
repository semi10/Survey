package com.fsdragon.android.survey;

import android.app.Activity;
import android.content.Context;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Semion on 07/07/2015.
 */
public class createSurvey extends Activity implements View.OnClickListener {
    List<String> questionList = new ArrayList<String>();
    Button next, previous, create;
    TextView progress;
    EditText question;
    int currentQuestionID = 0;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        setContentView(R.layout.new_survey);
        super.onCreate(savedInstanceState);

        next = (Button) findViewById(R.id.next_bt);
        previous = (Button) findViewById(R.id.previous_bt);
        create = (Button) findViewById(R.id.create_bt);
        progress = (TextView) findViewById(R.id.progress_tv);
        question = (EditText) findViewById(R.id.question_et);

        progress.setText(1 + "/" + 1);

        next.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                if (question.getText().toString().isEmpty())
                    Toast.makeText(createSurvey.this, "Enter your question", Toast.LENGTH_LONG).show();
                else {
                    currentQuestionID++;


                    if (questionList.size() < currentQuestionID) {
                        questionList.add(question.getText().toString());
                        progress.setText(currentQuestionID + 1 + "/" + (questionList.size() + 1));
                        question.setText("");
                    } else if (questionList.size() == currentQuestionID) {
                        questionList.set(currentQuestionID - 1, question.getText().toString());
                        progress.setText(currentQuestionID + 1 + "/" + (questionList.size() + 1));
                        question.setText("");
                    } else {
                        questionList.set(currentQuestionID - 1, question.getText().toString());
                        progress.setText(currentQuestionID + 1 + "/" + (questionList.size()));
                        question.setText(questionList.get(currentQuestionID));
                    }
                }
            }
        });


        previous.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                if (currentQuestionID == 0) return;
                    if (!question.getText().toString().isEmpty()){
                        if (questionList.size() ==  currentQuestionID) questionList.add(question.getText().toString());
                        else questionList.set(currentQuestionID, question.getText().toString());
                    }
                currentQuestionID--;
                question.setText(questionList.get(currentQuestionID));
                progress.setText(currentQuestionID + 1 + "/" + (questionList.size()));
            }
        });

        create.setOnClickListener(new View.OnClickListener() {
            public void onClick(View v) {
                if (!question.getText().toString().isEmpty()){
                    if (questionList.size() ==  currentQuestionID) questionList.add(question.getText().toString());
                    else questionList.set(currentQuestionID, question.getText().toString());
                }
                showDialog();
            }
        });
    }

    public void create(String sheetName) {
        saveExcelFile(this, sheetName, questionList);
    }


    private static boolean saveExcelFile(Context context, String sheetName, List<String> questionList){

        boolean success = false;

        //check if available and not read only
        if (!storageAvailable()) return false;

        File docsFolder = new File(Environment.getExternalStorageDirectory() + "/Documents");
        docsFolder.mkdirs();
        File file = new File(docsFolder, "Survey.xls");

        if (file.exists()){
            try {
                //Creating Input Stream
                FileInputStream myInput = new FileInputStream(file);

                //Create a POIFFileSystem object
                POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

                //Create a workbook using the File System
                HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

                HSSFSheet mySheet = myWorkBook.createSheet(sheetName);

                Cell cell = null;
                Row row = null;

                row = mySheet.createRow(0);
                //cell = row.createCell(0);
                mySheet.setColumnWidth(0, 8000);

                for(int i = 1; i <= questionList.size(); i++){
                    row = mySheet.createRow(i);
                    row = mySheet.getRow(i);
                    cell = row.createCell(0);
                    cell.setCellValue(questionList.get(i - 1).toString());
                }

                FileOutputStream os = null;

                try{
                    os = new FileOutputStream(file);
                    myWorkBook.write(os);
                    Log.w("FileUtils", "Writing file " + file);
                    os.close();
                    success = true;
                } catch (IOException e){
                    os.close();
                    Log.w("FileUtils", "Error writing " + file, e);
                } catch (Exception e){
                    os.close();
                    Log.w("FileUtils", "Failed to save file", e);
                } finally {
                    try{
                        if (null != os)
                            os.close();
                    }   catch (Exception ex){}
                }

            }catch (Exception e){e.printStackTrace();}
        }
        else{   //If no file found
            //New Workbook
            Workbook myWorkBook = new HSSFWorkbook();

            Cell c = null;


            //New Sheet
            Sheet mySheet = null;
            mySheet = myWorkBook.createSheet(sheetName);

            //Generate column headings
            Row row = null;
            Cell cell = null;

            row = mySheet.createRow(0);
            //cell = row.createCell(0);
            mySheet.setColumnWidth(0, 8000);

            for(int i = 1; i <= questionList.size(); i++){
                row = mySheet.createRow(i);
                row = mySheet.getRow(i);
                cell = row.createCell(0);
                cell.setCellValue(questionList.get(i - 1).toString());
            }

            FileOutputStream os = null;

            try{
                os = new FileOutputStream(file);
                myWorkBook.write(os);
                Log.w("FileUtils", "Writing file " + file);
                success = true;
            } catch (IOException e){
                Log.w("FileUtils", "Error writing " + file, e);
            } catch (Exception e){
                Log.w("FileUtils", "Failed to save file", e);
            }
            finally {
                try {
                    if (null != os)
                        os.close();
                }   catch (Exception ex){}
            }
        }
        return success;
    }

    private void showDialog(){
        NameIt nameIt = new NameIt();
        nameIt.show(getFragmentManager(), "Name It");
    }

    public static boolean storageAvailable(){
        //check if available and not read only
        String SD_state = Environment.getExternalStorageState();
        if (!Environment.MEDIA_MOUNTED.equals(SD_state) || Environment.MEDIA_MOUNTED_READ_ONLY.equals(SD_state)){
            Log.e("Storage", "Storage not available or read only");
            return false;
        }
        else return true;
    }

    @Override
    public void onClick(View v) {

    }
}