package com.example.wordtopdfconverter;

import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;

import android.annotation.TargetApi;
import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.github.barteksc.pdfviewer.PDFView;
import com.google.android.material.floatingactionbutton.FloatingActionButton;

import android.os.Environment;
import android.view.Menu;
import android.view.MenuInflater;
import android.view.MenuItem;
import android.view.View;
import android.widget.TextView;
import android.widget.Toast;

@TargetApi(Build.VERSION_CODES.FROYO)
// ... (existing imports)

public class MainActivity extends AppCompatActivity {

    private static final int PICK_PDF_FILE = 2;
    private final String storageDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS) + File.separator;
    private String outputPDF = storageDir + "Converted_PDF.pdf";
   // private TextView textView = null;
    private Uri document = null;
    private int fileCount = 0;
    private PDFView pdfView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        // apply the license if you have the Aspose.Words license...
        applyLicense();
        // get treeview and set its text
       // textView = findViewById(R.id.textView);
        //textView.setText("Select a Word DOCX file...");
        pdfView = findViewById(R.id.pdfView);
        // define click listener of floating button
        FloatingActionButton myFab = findViewById(R.id.fab);
        myFab.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                try {
                    // open Word file from file picker
                    openAndConvertFile();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });

    }
    private void openAndConvertFile() {
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        String[] mimetypes = {"application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword"};
        intent.setType("*/*");
        intent.putExtra(Intent.EXTRA_MIME_TYPES, mimetypes);
        startActivityForResult(intent, PICK_PDF_FILE);
    }


    @RequiresApi(api = Build.VERSION_CODES.KITKAT)

    @Override
    public void onActivityResult(int requestCode, int resultCode, Intent intent) {
        super.onActivityResult(requestCode, resultCode, intent);

        if (resultCode == Activity.RESULT_OK && intent != null) {
            document = intent.getData();

            try (InputStream inputStream = getContentResolver().openInputStream(document)) {
                Document doc = new Document(inputStream);
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                doc.save(outputStream, com.aspose.words.SaveFormat.PDF);
                pdfView.fromBytes(outputStream.toByteArray()).load();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
                Toast.makeText(this, "File not found: " + e.getMessage(), Toast.LENGTH_LONG).show();
            } catch (IOException e) {
                e.printStackTrace();
                Toast.makeText(this, e.getMessage(), Toast.LENGTH_LONG).show();
            } catch (Exception e) {
                e.printStackTrace();
                Toast.makeText(this, e.getMessage(), Toast.LENGTH_LONG).show();
            }
            //textView.setText("Tap here to download the PDF");
        }
    }


    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        MenuInflater inflater = getMenuInflater();
        inflater.inflate(R.menu.main_menu, menu);
        return true;
    }

    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        switch (item.getItemId()) {
            case R.id.action_download_pdf:
                // check if a document is selected before initiating the download
                if (document != null) {
                    convertAndDownload();
                } else {
                    Toast.makeText(this, "Please select a Word document first.", Toast.LENGTH_SHORT).show();
                }
                return true;
            default:
                return super.onOptionsItemSelected(item);
        }
    }

    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    private void convertAndDownload() {
        // open the selected document into an InputStream
        try (InputStream inputStream = getContentResolver().openInputStream(document)) {
            Document doc = new Document(inputStream);
            // update outputPDF with a numbered filename
            outputPDF = getNextFileName();
            // save DOCX as PDF
            doc.save(outputPDF);
            // show PDF file location in toast
            Toast.makeText(this, "File converted and saved at: " + outputPDF, Toast.LENGTH_LONG).show();
            // clear the stored document Uri
            document = null;
            // reset the text view
            //textView.setText("Select a Word DOCX file...");
            pdfView.fromFile(new File("")).load();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            Toast.makeText(this, "File not found: " + e.getMessage(), Toast.LENGTH_LONG).show();
        } catch (IOException e) {
            e.printStackTrace();
            Toast.makeText(this, e.getMessage(), Toast.LENGTH_LONG).show();
        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(this, e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }
    private void viewPDFFile() {
        // load PDF into the PDFView
        PDFView pdfView = findViewById(R.id.pdfView);
        pdfView.fromFile(new File(outputPDF)).load();
    }

    private void applyLicense() {
        // set license
        // License lic= new License();
        // add a raw resource directory within res and then add your license file as a resource...
        // InputStream inputStream = getResources().openRawResource(R.raw.license);
        try {
            // lic.setLicense(inputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    private String getNextFileName() {
        // Increment the fileCount and generate the new filename
        fileCount++;
        return storageDir + "Converted_PDF_" + fileCount + ".pdf";
    }
}
