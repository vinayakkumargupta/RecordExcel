package com.example.recordexcel;


import androidx.appcompat.app.AppCompatActivity;

import android.annotation.SuppressLint;
import android.os.Bundle;

import android.Manifest;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ImageView;
import android.widget.Toast;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.core.content.FileProvider;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class MainActivity extends AppCompatActivity {

    private EditText buyerEditText, styleEditText, fabricEditText, patternEditText;
    private ImageView dressImageView;
    private Button takePhotoButton, submitButton;

    private static final int REQUEST_CAMERA_PERMISSION = 100;
    private static final int REQUEST_IMAGE_CAPTURE = 1;

    private String currentPhotoPath;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        buyerEditText = findViewById(R.id.buyer_edittext);
        styleEditText = findViewById(R.id.style_edittext);
        fabricEditText = findViewById(R.id.fabric_edittext);
        patternEditText = findViewById(R.id.pattern_edittext);
        dressImageView = findViewById(R.id.dress_imageview);
        takePhotoButton = findViewById(R.id.take_photo_button);
        submitButton = findViewById(R.id.submit_button);

        takePhotoButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (ContextCompat.checkSelfPermission(MainActivity.this, Manifest.permission.CAMERA) != PackageManager.PERMISSION_GRANTED) {
                    ActivityCompat.requestPermissions(MainActivity.this, new String[] {Manifest.permission.CAMERA}, REQUEST_CAMERA_PERMISSION);
                } else {
                    dispatchTakePictureIntent();
                }
            }
        });

        submitButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                saveDataToExcel();
            }
        });
    }

    private void dispatchTakePictureIntent() {
        Intent takePictureIntent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
        if (takePictureIntent.resolveActivity(getPackageManager()) != null) {
            File photoFile = null;
            try {
                photoFile = createImageFile();
            } catch (IOException ex) {
                Toast.makeText(this, "Error occurred while creating the file", Toast.LENGTH_SHORT).show();
            }
            if (photoFile != null) {
                Uri photoURI = FileProvider.getUriForFile(this, "com.example.android.fileprovider", photoFile);
                takePictureIntent.putExtra(MediaStore.EXTRA_OUTPUT, photoURI);
                startActivityForResult(takePictureIntent, REQUEST_IMAGE_CAPTURE);
            }
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == REQUEST_IMAGE_CAPTURE && resultCode == RESULT_OK) {
            dressImageView.setImageURI(Uri.parse(currentPhotoPath));
        }
    }

    private File createImageFile() throws IOException {
        String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault()).format(new Date());
        String imageFileName = "JPEG_" + timeStamp + "_";
        File storageDir = getExternalFilesDir(Environment.DIRECTORY_PICTURES);
        File image = File.createTempFile(
                imageFileName,
                ".jpg",
                storageDir
        );
        currentPhotoPath = image.getAbsolutePath();
        return image;
    }

    private void saveDataToExcel() {
        // Create directory if it doesn't exist
        File directory = new File(Environment.getExternalStorageDirectory(), "Dress Data");
        if (!directory.exists()) {
            directory.mkdirs();
        }

        // Create new Excel file with unique name
        WorkbookSettings wbSettings = new WorkbookSettings();
        wbSettings.setLocale(new Locale("en", "EN"));
        String fileName = "DressData_" + new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault()).format(new Date()) + ".xls";
        File file = new File(directory, fileName);

        try {
            // Create new workbook and sheet
            WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
            WritableSheet sheet = workbook.createSheet("Sheet1", 0);

            // Add column labels to sheet
            Label label1 = new Label(0, 0, "Buyer");
            Label label2 = new Label(1, 0, "Style Number");
            Label label3 = new Label(2, 0, "Fabric");
            Label label4 = new Label(3, 0, "Pattern Number");
            sheet.addCell(label1);
            sheet.addCell(label2);
            sheet.addCell(label3);
            sheet.addCell(label4);

            // Retrieve data from EditText fields
            String buyer = buyerEditText.getText().toString().trim();
            String styleNumber = styleEditText.getText().toString().trim();
            String fabric = fabricEditText.getText().toString().trim();
            String patternNumber = patternEditText.getText().toString().trim();

            // Add data to sheet
            Label data1 = new Label(0, 1, buyer);
            Label data2 = new Label(1, 1, styleNumber);
            Label data3 = new Label(2, 1, fabric);
            Label data4 = new Label(3, 1, patternNumber);
            sheet.addCell(data1);
            sheet.addCell(data2);
            sheet.addCell(data3);
            sheet.addCell(data4);

            // Write and close workbook
            workbook.write();
            workbook.close();

            Toast.makeText(this, "Data saved to Excel successfully!", Toast.LENGTH_SHORT).show();

        } catch (IOException | WriteException e) {
            Toast.makeText(this, "Error occurred while saving data to Excel", Toast.LENGTH_SHORT).show();
            e.printStackTrace();
        }
    }
}
