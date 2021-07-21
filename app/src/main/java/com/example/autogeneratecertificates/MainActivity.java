package com.example.autogeneratecertificates;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ImageButton;

@SuppressWarnings("deprecation")
public class MainActivity extends AppCompatActivity {
    ImageButton templateone;
    Button browse;
    EditText entries;
    String path_xlsx;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        templateone = findViewById( R.id.templateone );
        browse=findViewById(R.id.browse);
        entries=findViewById(R.id.entries);
        getFilesDir();

        templateone.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent i1 = new Intent(getApplicationContext(), TemplateActivity.class );
                i1.putExtra( "template", "templateone" );
                i1.putExtra("path", path_xlsx);
                i1.putExtra("entries",entries.getText().toString());
                startActivity(i1);
            }
        } );

        browse.setOnClickListener(new View.OnClickListener() {
                                      @Override
                                      public void onClick(View v) {
                                          getPath();
                                      }
                                  }
        );
    }

    protected void getPath() {
        Intent i = new Intent(Intent.ACTION_GET_CONTENT);
        i.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        startActivityForResult(i, 1);
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        switch (requestCode) {
            case 1:
                if (resultCode == RESULT_OK) {
                    path_xlsx = data.getData().toString();
                }
                break;
        }
    }

}