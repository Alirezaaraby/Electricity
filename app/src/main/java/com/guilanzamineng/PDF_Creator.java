package com.guilanzamineng;

import android.content.Intent;
import android.os.Bundle;
import android.os.Environment;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.ViewTreeObserver;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.ScrollView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;

import com.google.android.material.floatingactionbutton.FloatingActionButton;
import com.google.android.material.snackbar.Snackbar;
import com.google.android.material.textfield.TextInputEditText;
import com.google.android.material.textfield.TextInputLayout;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.GrayColor;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.FileOutputStream;
import java.text.DecimalFormat;

public class PDF_Creator extends AppCompatActivity {

    FloatingActionButton fab_PDF;

    TextInputLayout SUBSTATION_Text_Input_Layout;
    TextInputLayout BAY_NO_Text_Input_Layout;
    TextInputLayout PANEL_NU_Text_Input_Layout;
    TextInputLayout METER_MANUFACTURER_Text_Input_Layout;
    TextInputLayout CT_RATIO_Text_Input_Layout;
    TextInputLayout PT_RATIO_Text_Input_Layout;
    TextInputLayout CT_RATIO_CLASS__Text_Input_Layout;
    TextInputLayout PT_RATIO_CLASS__Text_Input_Layout;
    TextInputLayout METER_SCALE_RANGE_Text_Input_Layout;
    TextInputLayout METER_MODEL__SERIAL_NO_Text_Input_Layout;
    TextInputLayout REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout;
    TextInputLayout REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout;
    TextInputLayout REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout;
    TextInputLayout I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout;
    TextInputLayout I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout;
    TextInputLayout I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout;
    TextInputLayout REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout;
    TextInputLayout REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout;
    TextInputLayout REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout;
    TextInputLayout V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout;
    TextInputLayout V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout;
    TextInputLayout V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout;
    TextInputLayout ACTIVE_POWER;
    TextInputLayout REACTIVE_POWER;
    TextInputLayout COS_PHI;
    TextInputLayout COMMENT;
    TextInputLayout TEST_BY;
    TextInputLayout DATE;

    TextInputEditText SUBSTATION_Text_Input_Layout_;
    TextInputEditText BAY_NO_Text_Input_Layout_;
    TextInputEditText PANEL_NU_Text_Input_Layout_;
    TextInputEditText METER_MANUFACTURER_Text_Input_Layout_;
    TextInputEditText CT_RATIO_Text_Input_Layout_;
    TextInputEditText PT_RATIO_Text_Input_Layout_;
    TextInputEditText CT_RATIO_CLASS__Text_Input_Layout_;
    TextInputEditText PT_RATIO_CLASS__Text_Input_Layout_;
    TextInputEditText METER_SCALE_RANGE_Text_Input_Layout_;
    TextInputEditText METER_MODEL__SERIAL_NO_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout_;
    TextInputEditText I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout_;
    TextInputEditText I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout_;
    TextInputEditText I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout_;
    TextInputEditText REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout_;
    TextInputEditText V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout_;
    TextInputEditText V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout_;
    TextInputEditText V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout_;
    TextInputEditText ACTIVE_POWER_;
    TextInputEditText REACTIVE_POWER_;
    TextInputEditText COS_PHI_;
    TextInputEditText COMMENT_;
    TextInputEditText TEST_BY_;
    TextInputEditText DATE_;

    FloatingActionButton scrolltotop;
    ScrollView scrollViewEventDetails;

    MenuItem visibility;

    private static final String TAG = "APP";

    CheckBox ch;

    private Menu menu;


    @Override
    public boolean onCreateOptionsMenu(final Menu menu) {
        getMenuInflater().inflate(R.menu.excel_menu, menu);
        menu.findItem(R.id.Show_Excel_Menu_Visible_Fab).setVisible(false);
        this.menu = menu;
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(@NonNull final MenuItem item) {
        int id = item.getItemId();
        if (id == R.id.Show_Excel_Menu_Visible_Fab) {
            fab_PDF.setVisibility(View.VISIBLE);
            item.setVisible(false);
        }
        else{
            Intent intent = new Intent(PDF_Creator.this , MainActivity.class);
            startActivity(intent);
            finish();
        }
        return true;
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_pdf_creator);

        fab_PDF = (FloatingActionButton) findViewById(R.id.clear_all_pdf);

        try {
            getSupportActionBar().setDisplayHomeAsUpEnabled(true);
            getSupportActionBar().setDisplayShowHomeEnabled(true);
        }
        catch (Exception e){
            Toast.makeText(this, "An Undefined Error Occurred", Toast.LENGTH_SHORT).show();
        }

        SUBSTATION_Text_Input_Layout = findViewById(R.id.SUBSTATION_PDF);
        BAY_NO_Text_Input_Layout = findViewById(R.id.BAY_NO_PDF);
        PANEL_NU_Text_Input_Layout = findViewById(R.id.PANEL_NO_PDF);
        METER_MANUFACTURER_Text_Input_Layout = findViewById(R.id.METER_MANUFACTURER_PDF);
        CT_RATIO_Text_Input_Layout = findViewById(R.id.CT_RATIO_PDF);
        CT_RATIO_CLASS__Text_Input_Layout = findViewById(R.id.CT_RATIO_CLASS_PDF);
        PT_RATIO_Text_Input_Layout = findViewById(R.id.PT_RATIO_PDF);
        PT_RATIO_CLASS__Text_Input_Layout = findViewById(R.id.PT_RATIO_CLASS_PDF);
        METER_SCALE_RANGE_Text_Input_Layout = findViewById(R.id.METER_SCALE_RANGE_PDF);
        METER_MODEL__SERIAL_NO_Text_Input_Layout = findViewById(R.id.METER_MODEL__SERIAL_NO_PDF);
        REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout = findViewById(R.id.REAL_VALUEI_LOAD_PHASE_A_PDF);
        REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout = findViewById(R.id.REAL_VALUEI_LOAD_PH_B_PDF);
        REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout = findViewById(R.id.REAL_VALUEI_LOAD_PH_C_PDF);
        I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout = findViewById(R.id.I_ONLOAD_by_clamp_PH_A_PDF);
        I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout = findViewById(R.id.I_ONLOAD_by_clamp_PH_B_PDF);
        I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout = findViewById(R.id.I_ONLOAD_by_clamp_PH_c__PDF);
        REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout = findViewById(R.id.REAL_VALUE_V_LOAD_PH_A_PDF);
        REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout = findViewById(R.id.REAL_VALUE_V_LOAD_PH_B_PDF);
        REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout = findViewById(R.id.REAL_VALUE_V_LOAD_PH_C_PDF);
        V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_A_PDF);
        V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_B_PDF);
        V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_C_PDF);
        ACTIVE_POWER = findViewById(R.id.ACTIVE_POWER_PDF);
        REACTIVE_POWER = findViewById(R.id.REACTIVE_POWER_PDF);
        COS_PHI = findViewById(R.id.COS_PHI_PDF);
        COMMENT = findViewById(R.id.COMMENT_PDF);
        TEST_BY = findViewById(R.id.TEST_BY_PDF);
        DATE = findViewById(R.id.DATE_PDF);

        SUBSTATION_Text_Input_Layout_ = findViewById(R.id.SUBSTATION______PDF);
        BAY_NO_Text_Input_Layout_ = findViewById(R.id.BAY_NO______PDF);
        PANEL_NU_Text_Input_Layout_ = findViewById(R.id.PANEL_NO______PDF);
        METER_MANUFACTURER_Text_Input_Layout_ = findViewById(R.id.METER_MANUFACTURER______PDF);
        CT_RATIO_Text_Input_Layout_ = findViewById(R.id.CT_RATIO______PDF);
        CT_RATIO_CLASS__Text_Input_Layout_ = findViewById(R.id.CT_RATIO_CLASS______PDF);
        PT_RATIO_Text_Input_Layout_ = findViewById(R.id.PT_RATIO______PDF);
        PT_RATIO_CLASS__Text_Input_Layout_ = findViewById(R.id.PT_RATIO_CLASS______PDF);
        METER_SCALE_RANGE_Text_Input_Layout_ = findViewById(R.id.METER_SCALE_RANGE_Text_Input_Edit_Text_PDF);
        METER_MODEL__SERIAL_NO_Text_Input_Layout_ = findViewById(R.id.METER_MODEL__SERIAL__NO_PDF);
        REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout_ = findViewById(R.id.REAL_VALUEI_LOAD_PHASE_A__PDF);
        REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout_ = findViewById(R.id.REAL_VALUEI_LOAD_PH_B__PDF);
        REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout_ = findViewById(R.id.REAL_VALUEI_LOAD_PH_C__PDF);
        I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout_ = findViewById(R.id.I_ONLOAD_by_clamp_PH_A__PDF);
        I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout_ = findViewById(R.id.I_ONLOAD_by_clamp_PH_B__PDF);
        I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout_ = findViewById(R.id.I_ONLOAD_by_clamp_PH_C_PDF);
        REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout_ = findViewById(R.id.REAL_VALUE_V_LOAD_PH_A__PDF);
        REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout_ = findViewById(R.id.REAL_VALUE_V_LOAD_PH_B__PDF);
        REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout_ = findViewById(R.id.REAL_VALUE_V_LOAD_PH_C__PDF);
        V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout_ = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_A__PDF);
        V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout_ = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_B__PDF);
        V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout_ = findViewById(R.id.V_ONLOAD_BY_V_M_V_LOAD_PH_C__PDF);
        ACTIVE_POWER_ = findViewById(R.id.ACTIVE_POWER__PDF);
        REACTIVE_POWER_ = findViewById(R.id.REACTIVE_POWER__PDF);
        COS_PHI_ = findViewById(R.id.COS_PHI__PDF);
        COMMENT_ = findViewById(R.id.COMMENT__PDF);
        TEST_BY_ = findViewById(R.id.TEST_BY__PDF);
        DATE_ = findViewById(R.id.DATE__PDF);

        scrollViewEventDetails = findViewById(R.id.scrollView_PDF);

        scrolltotop = findViewById(R.id.top_PDF);
        scrollViewEventDetails.getViewTreeObserver().addOnScrollChangedListener(new ViewTreeObserver.OnScrollChangedListener() {
            @Override
            public void onScrollChanged() {
                if (4000 >= scrollViewEventDetails.getScrollY()){
                    scrolltotop.setVisibility(View.VISIBLE);
                }
                if (4000 <= scrollViewEventDetails.getScrollY()){
                    scrolltotop.setVisibility(View.INVISIBLE);
                }
            }
        });

        scrolltotop.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                if (4000 >= scrollViewEventDetails.getScrollY()){
                    scrollViewEventDetails.fullScroll(View.FOCUS_DOWN);
                }
            }
        });


        fab_PDF.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                SUBSTATION_Text_Input_Layout_.setText("");
                BAY_NO_Text_Input_Layout_.setText("");
                PANEL_NU_Text_Input_Layout_.setText("");
                METER_MANUFACTURER_Text_Input_Layout_.setText("");
                CT_RATIO_Text_Input_Layout_.setText("");
                CT_RATIO_CLASS__Text_Input_Layout_.setText("");
                PT_RATIO_Text_Input_Layout_.setText("");
                PT_RATIO_CLASS__Text_Input_Layout_.setText("");
                METER_SCALE_RANGE_Text_Input_Layout_.setText("");
                METER_MODEL__SERIAL_NO_Text_Input_Layout_.setText("");
                REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout_.setText("");
                REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout_.setText("");
                REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout_.setText("");
                I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout_.setText("");
                I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout_.setText("");
                I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout_.setText("");
                REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout_.setText("");
                REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout_.setText("");
                REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout_.setText("");
                V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout_.setText("");
                V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout_.setText("");
                V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout_.setText("");
                ACTIVE_POWER_.setText("");
                REACTIVE_POWER_.setText("");
                COS_PHI_.setText("");
                COMMENT_.setText("");
                TEST_BY_.setText("");
                DATE_.setText("");
                Snackbar.make(view, "Edit Texts Are Empty Now", Snackbar.LENGTH_SHORT)
                        .setAction("Action", null).show();
            }
        });
        fab_PDF.setOnLongClickListener(new View.OnLongClickListener() {
            @Override
            public boolean onLongClick(View v) {
                fab_PDF.setVisibility(View.GONE);

                menu.findItem(R.id.Show_Excel_Menu_Visible_Fab).setVisible(true);

                Snackbar.make(v, "Clear Button Is Invisible Now", Snackbar.LENGTH_LONG)
                        .setAction("Action", null).show();

                return true;
            }
        });

        Button Save = findViewById(R.id.outlinedButton_PDF);
        Save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveTheFile();
            }
        });
    }

    public void createPDF() {
        try {
            if (!CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !SUBSTATION_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !BAY_NO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty() && !DATE.getEditText().getText().toString().trim().isEmpty()){
                if(CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().contains("/") && PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().contains("/")) {
                    String[] CT = CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");
                    String[] PT = PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().split("/");
                        if (Integer.parseInt(CT[0]) == (int) Integer.parseInt(CT[0]) && Integer.parseInt(CT[1]) == (int) Integer.parseInt(CT[1]) && Integer.parseInt(PT[0]) == (int) Integer.parseInt(PT[0]) && Integer.parseInt(PT[1]) == (int) Integer.parseInt(PT[1])) {
                            //
                            Document document = new Document();
                            String mFileName = SUBSTATION_Text_Input_Layout.getEditText().getText().toString() + "_" + BAY_NO_Text_Input_Layout.getEditText().getText().toString() +"_"+ "metter" +"_"+ DATE.getEditText().getText().toString();
                            String mFilePath = Environment.getExternalStorageDirectory().getPath() + "/Electricity/PDF/"+ mFileName + ".pdf";

                            int J3 = Integer.parseInt(CT[0]) / Integer.parseInt(CT[1]);
                            int J4 = Integer.parseInt(PT[0]) / Integer.parseInt(PT[1]);
                            // Error 1
                            double B15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString().trim());
                            double B11 = Double.parseDouble(REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D11 = Double.parseDouble(REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
                            double F15 = Double.parseDouble(I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double F11 = Double.parseDouble(REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            //Error 2
                            double F28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double F25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim());
                            double D25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            double B28 = Double.parseDouble(V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim());
                            double B25 = Double.parseDouble(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim());
                            //
                            double error_2_1 = Math.abs(((B28*J4-B25)/B25))*100;
                            double error_2_2 = Math.abs((((D28)*J4)-D25)/D25)*100;
                            double error_2_3 = Math.abs((((F28)*J4)-F25)/F25)*100;
                            //
                            double error_1_1 = Math.abs(((B15*J3-B11)/B11))*100;
                            double error_1_2 = Math.abs((((D15)*J3)-D11)/D11)*100;
                            double error_1_3 = Math.abs((((F15)*J3)-F11)/F11)*100;
                            try {
                                PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(mFilePath));
                                document.open();

                                PdfPTable table = new PdfPTable(8);
                                table.setWidthPercentage(100);
                                table.getDefaultCell().setUseAscender(true);
                                table.getDefaultCell().setUseDescender(true);

                                float[] columnWidths = new float[]{25f, 8f, 10f, 10f, 17f, 10f, 10f,10f};
                                table.setWidths(columnWidths);

                                Font f = new Font(Font.FontFamily.HELVETICA, 10, Font.NORMAL, GrayColor.BLACK);
                                Font f_1 = new Font(Font.FontFamily.HELVETICA, 8, Font.NORMAL, GrayColor.BLACK);
                                PdfPCell cell1 = new PdfPCell(new Phrase("GUILAN REGIONAL ELECTRIC COMPANY", f));
                                cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell1.setColspan(8);
                                cell1.setPaddingBottom(4);
                                cell1.setBorderWidth(2);
                                table.addCell(cell1);

                                table.setTableEvent(new BorderEvent());


                                PdfPCell cell2 = new PdfPCell(new Phrase("METERING TEST SHEET", f));
                                cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell2.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell2.setColspan(8);
                                cell2.setPaddingBottom(4);
                                cell2.setBorderWidth(2);
                                table.addCell(cell2);

                                //Row 3
                                PdfPCell cell3 = new PdfPCell(new Phrase("SUBSTATION",f));
                                cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell3.setColspan(2);
                                cell3.setBorderWidth(2);
                                cell3.setFixedHeight(30f);
                                table.addCell(cell3);

                                PdfPCell cell4 = new PdfPCell(new Phrase(SUBSTATION_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell4.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell4.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell4.setColspan(2);
                                cell4.setBorderWidth(2);
                                table.addCell(cell4);

                                PdfPCell cell5 = new PdfPCell(new Phrase("CT RATIO",f));
                                cell5.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell5.setColspan(1);
                                cell5.setBorderWidth(2);
                                table.addCell(cell5);

                                PdfPCell cell6 = new PdfPCell(new Phrase(CT_RATIO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell6.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell6.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell6.setColspan(1);
                                cell6.setBorderWidth(2);
                                table.addCell(cell6);

                                PdfPCell cell7 = new PdfPCell(new Phrase("CLASS",f));
                                cell7.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell7.setColspan(1);
                                cell7.setBorderWidth(2);
                                table.addCell(cell7);

                                PdfPCell cell8 = new PdfPCell(new Phrase(CT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString(),f));
                                cell8.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell8.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell8.setColspan(1);
                                cell8.setBorderWidth(2);
                                table.addCell(cell8);

                                //Row 4
                                PdfPCell cell9 = new PdfPCell(new Phrase("BAY NO",f));
                                cell9.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell9.setColspan(2);
                                cell9.setBorderWidth(2);
                                cell9.setFixedHeight(30f);
                                table.addCell(cell9);

                                PdfPCell cell10 = new PdfPCell(new Phrase(BAY_NO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell10.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell10.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell10.setColspan(2);
                                cell10.setBorderWidth(2);
                                table.addCell(cell10);

                                PdfPCell cell11 = new PdfPCell(new Phrase("PT RATIO",f));
                                cell11.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell11.setBorderWidth(2);
                                table.addCell(cell11);

                                PdfPCell cell12 = new PdfPCell(new Phrase(PT_RATIO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell12.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell12.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell12.setBorderWidth(2);
                                table.addCell(cell12);

                                PdfPCell cell13 = new PdfPCell(new Phrase("CLASS",f));
                                cell13.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell13.setBorderWidth(2);
                                table.addCell(cell13);

                                PdfPCell cell14 = new PdfPCell(new Phrase(PT_RATIO_CLASS__Text_Input_Layout.getEditText().getText().toString(),f));
                                cell14.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell14.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell14.setBorderWidth(2);
                                table.addCell(cell14);

                                //Row 5
                                PdfPCell cell15 = new PdfPCell(new Phrase("PANEL NO",f));
                                cell15.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell15.setBorderWidth(2);
                                cell15.setColspan(2);
                                cell15.setFixedHeight(30f);
                                table.addCell(cell15);

                                PdfPCell cell16 = new PdfPCell(new Phrase(PANEL_NU_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell16.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell16.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell16.setBorderWidth(2);
                                cell16.setColspan(2);
                                table.addCell(cell16);

                                PdfPCell cell17 = new PdfPCell(new Phrase("METER SCALE RANGE",f));
                                cell17.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell17.setBorderWidth(2);
                                cell17.setColspan(2);
                                table.addCell(cell17);

                                PdfPCell cell18 = new PdfPCell(new Phrase(METER_SCALE_RANGE_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell18.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell18.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell18.setBorderWidth(2);
                                cell18.setColspan(2);
                                table.addCell(cell18);

                                //Row 5
                                PdfPCell cell19 = new PdfPCell(new Phrase("METER MANUFACTURER ",f));
                                cell19.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell19.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell19.setFixedHeight(30f);
                                cell19.setBorderWidth(2);
                                cell19.setColspan(2);
                                table.addCell(cell19);


                                PdfPCell cell20 = new PdfPCell(new Phrase(METER_MANUFACTURER_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell20.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell20.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell20.setBorderWidth(2);
                                cell20.setColspan(2);
                                table.addCell(cell20);

                                PdfPCell cell21 = new PdfPCell(new Phrase("METER MODEL--SERIAL NO",f));
                                cell21.setBorderWidth(2);
                                cell21.setColspan(2);
                                cell21.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell21.setHorizontalAlignment(Element.ALIGN_CENTER);
                                table.addCell(cell21);

                                PdfPCell cell22 = new PdfPCell(new Phrase(METER_MODEL__SERIAL_NO_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell22.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell22.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell22.setBorderWidth(2);
                                cell22.setColspan(2);
                                table.addCell(cell22);

                                // Row 6
                                PdfPCell cell23 = new PdfPCell(new Phrase("REAL VALUE-I LOAD",f));
                                cell23.setFixedHeight(70f);
                                cell23.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell23.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell23.setColspan(1);
                                cell23.setRowspan(3);
                                table.addCell(cell23);

                                PdfPCell cell24 = new PdfPCell(new Phrase("PHASE A",f));
                                cell24.setFixedHeight(35f);
                                cell24.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell24.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell24.setColspan(2);
                                cell24.setRowspan(1);
                                table.addCell(cell24);

                                PdfPCell cell25 = new PdfPCell(new Phrase("PH.B",f));
                                cell25.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell25.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell25.setColspan(2);
                                cell25.setRowspan(1);
                                table.addCell(cell25);

                                PdfPCell cell26 = new PdfPCell(new Phrase("PH.C",f));
                                cell26.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell26.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell26.setColspan(3);
                                cell26.setRowspan(1);
                                table.addCell(cell26);
                                //
                                PdfPCell cell27 = new PdfPCell(new Phrase(REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell27.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell27.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell27.setColspan(2);
                                cell27.setRowspan(2);
                                table.addCell(cell27);

                                PdfPCell cell28 = new PdfPCell(new Phrase(REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell28.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell28.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell28.setColspan(2);
                                cell28.setRowspan(2);
                                table.addCell(cell28);

                                PdfPCell cell29 = new PdfPCell(new Phrase(REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell29.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell29.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell29.setColspan(3);
                                cell29.setRowspan(2);
                                table.addCell(cell29);
                                //Row 9
                                PdfPCell cell30 = new PdfPCell(new Phrase("I ONLOAD by clamp",f));
                                cell30.setFixedHeight(70f);
                                cell30.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell30.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell30.setColspan(1);
                                cell30.setRowspan(4);
                                table.addCell(cell30);

                                PdfPCell cell31 = new PdfPCell(new Phrase("PH.A",f));
                                cell31.setFixedHeight(35f);
                                cell31.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell31.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell31.setColspan(2);
                                cell31.setRowspan(2);
                                table.addCell(cell31);

                                PdfPCell cell32 = new PdfPCell(new Phrase("PH.B",f));
                                cell32.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell32.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell32.setColspan(2);
                                cell32.setRowspan(2);
                                table.addCell(cell32);

                                PdfPCell cell33 = new PdfPCell(new Phrase("PH.C",f));
                                cell33.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell33.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell33.setColspan(3);
                                cell33.setRowspan(2);
                                table.addCell(cell33);

                                //

                                PdfPCell cell34 = new PdfPCell(new Phrase(I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell34.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell34.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell34.setColspan(2);
                                cell34.setRowspan(2);
                                table.addCell(cell34);

                                PdfPCell cell35 = new PdfPCell(new Phrase(I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell35.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell35.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell35.setColspan(2);
                                cell35.setRowspan(2);
                                table.addCell(cell35);

                                PdfPCell cell36 = new PdfPCell(new Phrase(I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell36.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell36.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell36.setColspan(3);
                                cell36.setRowspan(2);
                                table.addCell(cell36);

                                //Row 13

                                PdfPCell cell37 = new PdfPCell(new Phrase("%ERROR",f));
                                cell37.setFixedHeight(35f);
                                cell37.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell37.setColspan(1);
                                cell37.setRowspan(2);
                                table.addCell(cell37);

                                PdfPCell cell38 = new PdfPCell(new Phrase(new DecimalFormat("##.###").format(error_1_1),f));
                                cell38.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell38.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell38.setColspan(2);
                                cell38.setRowspan(2);
                                table.addCell(cell38);

                                PdfPCell cell39 = new PdfPCell(new Phrase(new DecimalFormat("##.###").format(error_1_2),f));
                                cell39.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell39.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell39.setColspan(2);
                                cell39.setRowspan(2);
                                table.addCell(cell39);

                                PdfPCell cell40 = new PdfPCell(new Phrase(new DecimalFormat("##.###").format(error_1_3),f));
                                cell40.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell40.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell40.setColspan(3);
                                cell40.setRowspan(2);
                                table.addCell(cell40);

                                PdfPCell cell41 = new PdfPCell(new Phrase("%ERROR IN RANGE",f));
                                cell41.setFixedHeight(35f);
                                cell41.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell41.setColspan(1);
                                cell41.setRowspan(2);
                                table.addCell(cell41);

                                PdfPCell cell42 = new PdfPCell(new Phrase("",f));
                                cell42.setColspan(2);
                                cell42.setRowspan(2);
                                table.addCell(cell42);

                                PdfPCell cell43 = new PdfPCell(new Phrase("",f));
                                cell43.setColspan(2);
                                cell43.setRowspan(2);
                                table.addCell(cell43);

                                PdfPCell cell44 = new PdfPCell(new Phrase("",f));
                                cell44.setColspan(3);
                                cell44.setRowspan(2);
                                table.addCell(cell44);
                                //

                                PdfPCell cell45 = new PdfPCell(new Phrase("%ERROR OUT OF RANGE",f));
                                cell45.setFixedHeight(35f);
                                cell45.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell45.setColspan(1);
                                cell45.setRowspan(2);
                                table.addCell(cell45);

                                PdfPCell cell46 = new PdfPCell(new Phrase("",f));
                                cell46.setColspan(2);
                                cell46.setRowspan(2);
                                table.addCell(cell46);

                                PdfPCell cell47 = new PdfPCell(new Phrase("",f));
                                cell47.setColspan(2);
                                cell47.setRowspan(2);
                                table.addCell(cell47);

                                PdfPCell cell48 = new PdfPCell(new Phrase("",f));
                                cell48.setColspan(3);
                                cell48.setRowspan(2);
                                table.addCell(cell48);

                                // Row 6
                                PdfPCell cell49 = new PdfPCell(new Phrase("REAL VALUE-V LOAD",f));
                                cell49.setFixedHeight(70f);
                                cell49.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell49.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell49.setBorderWidthTop(2);
                                cell49.setColspan(1);
                                cell49.setRowspan(3);
                                table.addCell(cell49);

                                PdfPCell cell50 = new PdfPCell(new Phrase("PH.A",f));
                                cell50.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell50.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell50.setFixedHeight(35f);
                                cell50.setBorderWidthTop(2);
                                cell50.setColspan(2);
                                cell50.setRowspan(1);
                                table.addCell(cell50);

                                PdfPCell cell51 = new PdfPCell(new Phrase("PH.B",f));
                                cell51.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell51.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell51.setBorderWidthTop(2);
                                cell51.setColspan(2);
                                cell51.setRowspan(1);
                                table.addCell(cell51);

                                PdfPCell cell52 = new PdfPCell(new Phrase("PH.C",f));
                                cell52.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell52.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell52.setBorderWidthTop(2);
                                cell52.setColspan(3);
                                cell52.setRowspan(1);
                                table.addCell(cell52);
                                //
                                PdfPCell cell53 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell53.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell53.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell53.setColspan(2);
                                cell53.setRowspan(2);
                                table.addCell(cell53);

                                PdfPCell cell54 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell54.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell54.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell54.setColspan(2);
                                cell54.setRowspan(2);
                                table.addCell(cell54);

                                PdfPCell cell55 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell55.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell55.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell55.setColspan(3);
                                cell55.setRowspan(2);
                                table.addCell(cell55);
                                //Row 9
                                PdfPCell cell56 = new PdfPCell(new Phrase("V ONLOAD BY V.M-V LOAD",f_1));
                                cell56.setFixedHeight(70f);
                                cell56.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell56.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell56.setColspan(1);
                                cell56.setRowspan(4);
                                table.addCell(cell56);

                                PdfPCell cell57 = new PdfPCell(new Phrase("PH.A",f));
                                cell57.setFixedHeight(35f);
                                cell57.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell57.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell57.setColspan(2);
                                cell57.setRowspan(2);
                                table.addCell(cell57);

                                PdfPCell cell58 = new PdfPCell(new Phrase("PH.B",f));
                                cell58.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell58.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell58.setColspan(2);
                                cell58.setRowspan(2);
                                table.addCell(cell58);

                                PdfPCell cell59 = new PdfPCell(new Phrase("PH.C",f));
                                cell59.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell59.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell59.setColspan(3);
                                cell59.setRowspan(2);
                                table.addCell(cell59);

                                PdfPCell cell60 = new PdfPCell(new Phrase(V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell60.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell60.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell60.setColspan(2);
                                cell60.setRowspan(2);
                                table.addCell(cell60);

                                PdfPCell cell61 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell61.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell61.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell61.setColspan(2);
                                cell61.setRowspan(2);
                                table.addCell(cell61);

                                PdfPCell cell62 = new PdfPCell(new Phrase(REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString(),f));
                                cell62.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell62.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell62.setColspan(3);
                                cell62.setRowspan(2);
                                table.addCell(cell62);
                                //Row 13

                                PdfPCell cell63 = new PdfPCell(new Phrase("%ERROR",f));
                                cell63.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell63.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell63.setFixedHeight(35f);
                                cell63.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell63.setColspan(1);
                                cell63.setRowspan(2);
                                table.addCell(cell63);

                                PdfPCell cell64 = new PdfPCell(new Phrase(new DecimalFormat("##.###").format(error_2_1),f));
                                cell64.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell64.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell64.setColspan(2);
                                cell64.setRowspan(2);
                                table.addCell(cell64);

                                PdfPCell cell65 = new PdfPCell(new Phrase(new DecimalFormat("##.###").format(error_2_2),f));
                                cell65.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell65.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell65.setColspan(2);
                                cell65.setRowspan(2);
                                table.addCell(cell65);

                                PdfPCell cell66 = new PdfPCell(new Phrase(new DecimalFormat("##.###").format(error_2_3),f));
                                cell66.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell66.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell66.setColspan(3);
                                cell66.setRowspan(2);
                                table.addCell(cell66);

                                PdfPCell cell67 = new PdfPCell(new Phrase("%ERROR IN RANGE",f));
                                cell67.setFixedHeight(35f);
                                cell67.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell67.setColspan(1);
                                cell67.setRowspan(2);
                                table.addCell(cell67);

                                PdfPCell cell68 = new PdfPCell(new Phrase("",f));
                                cell68.setColspan(2);
                                cell68.setRowspan(2);
                                table.addCell(cell68);

                                PdfPCell cell69 = new PdfPCell(new Phrase("",f));
                                cell69.setColspan(2);
                                cell69.setRowspan(2);
                                table.addCell(cell69);

                                PdfPCell cell70 = new PdfPCell(new Phrase("",f));
                                cell70.setColspan(3);
                                cell70.setRowspan(2);
                                table.addCell(cell70);

                                //


                                PdfPCell cell71 = new PdfPCell(new Phrase("%ERROR OUT OF RANGE",f));
                                cell71.setFixedHeight(35f);
                                cell71.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell71.setColspan(1);
                                cell71.setRowspan(2);
                                table.addCell(cell71);

                                PdfPCell cell72 = new PdfPCell(new Phrase("",f));
                                cell72.setColspan(2);
                                cell72.setRowspan(2);
                                table.addCell(cell72);

                                PdfPCell cell73 = new PdfPCell(new Phrase("",f));
                                cell73.setColspan(2);
                                cell73.setRowspan(2);
                                table.addCell(cell73);

                                PdfPCell cell74 = new PdfPCell(new Phrase("",f));
                                cell74.setColspan(3);
                                cell74.setRowspan(2);
                                table.addCell(cell74);

                                //

                                PdfPCell cell75 = new PdfPCell(new Phrase("ACTIVE POWER",f));
                                cell75.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell75.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell75.setBorderWidth(2);
                                cell75.setColspan(1);
                                cell75.setRowspan(2);
                                table.addCell(cell75);

                                PdfPCell cell76 = new PdfPCell(new Phrase(ACTIVE_POWER.getEditText().getText().toString(),f));
                                cell76.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell76.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell76.setBorderWidth(2);
                                cell76.setColspan(2);
                                cell76.setRowspan(2);
                                table.addCell(cell76);

                                PdfPCell cell77 = new PdfPCell(new Phrase("REACTIVE POWER",f));
                                cell77.setFixedHeight(40f);
                                cell77.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell77.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell77.setBorderWidth(2);
                                cell77.setColspan(2);
                                cell77.setRowspan(2);
                                table.addCell(cell77);

                                PdfPCell cell78 = new PdfPCell(new Phrase(REACTIVE_POWER.getEditText().getText().toString(),f));
                                cell78.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell78.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell78.setBorderWidth(2);
                                cell78.setColspan(3);
                                cell78.setRowspan(2);
                                table.addCell(cell78);

                                //

                                PdfPCell cell79 = new PdfPCell(new Phrase("COS PHI",f));
                                cell79.setFixedHeight(40f);
                                cell79.setBorderWidth(2);
                                cell79.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell79.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell79.setColspan(1);
                                cell79.setRowspan(2);
                                table.addCell(cell79);

                                PdfPCell cell80 = new PdfPCell(new Phrase(COS_PHI.getEditText().getText().toString(),f));
                                cell80.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell80.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell80.setBorderWidth(2);
                                cell80.setColspan(2);
                                cell80.setRowspan(2);
                                table.addCell(cell80);

                                PdfPCell cell81 = new PdfPCell(new Phrase("COMMENT",f));
                                cell81.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell81.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell81.setBorderWidth(2);
                                cell81.setColspan(2);
                                cell81.setRowspan(2);
                                table.addCell(cell81);

                                PdfPCell cell82 = new PdfPCell(new Phrase(COMMENT.getEditText().getText().toString(),f));
                                cell82.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell82.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell82.setBorderWidth(2);
                                cell82.setColspan(3);
                                cell82.setRowspan(2);
                                table.addCell(cell82);

                                //

                                PdfPCell cell83 = new PdfPCell(new Phrase("TEST BY",f));
                                cell83.setFixedHeight(40f);
                                cell83.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell83.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell83.setBorderWidth(2);
                                cell83.setColspan(1);
                                cell83.setRowspan(3);
                                table.addCell(cell83);

                                PdfPCell cell84 = new PdfPCell(new Phrase(TEST_BY.getEditText().getText().toString(),f));
                                cell84.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell84.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell84.setBorderWidth(2);
                                cell84.setColspan(2);
                                cell84.setRowspan(3);
                                table.addCell(cell84);

                                PdfPCell cell85 = new PdfPCell(new Phrase("DATE",f));
                                cell85.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell85.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell85.setBorderWidth(2);
                                cell85.setColspan(2);
                                cell85.setRowspan(3);
                                table.addCell(cell85);

                                PdfPCell cell86 = new PdfPCell(new Phrase(DATE.getEditText().getText().toString(),f));
                                cell86.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                cell86.setHorizontalAlignment(Element.ALIGN_CENTER);
                                cell86.setBorderWidth(2);
                                cell86.setColspan(3);
                                cell86.setRowspan(3);
                                table.addCell(cell86);

                                HeaderFooterPageEvent event = new HeaderFooterPageEvent();
                                writer.setPageEvent(event);

                                document.add(table);
                                document.close();
                                Toast.makeText(PDF_Creator.this, "Saved To" + mFilePath, Toast.LENGTH_SHORT).show();
                            }
                            catch (Exception e) {
                                Toast.makeText(PDF_Creator.this, "Format Not Support For CT RATIO AND PT RATIO", Toast.LENGTH_SHORT).show();
                            }
                    }
                }
                else{
                    Toast.makeText(PDF_Creator.this, "Format Doesn't Support", Toast.LENGTH_SHORT).show();
                }
            }

            else{
//                if (SUBSTATION_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    SUBSTATION_Text_Input_Layout.setError("This Field Is Required");
//                if (BAY_NO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    BAY_NO_Text_Input_Layout.setError("This Field Is Required");
//                if (DATE.getEditText().getText().toString().trim().isEmpty())
//                    DATE.setError("This Field Is Required");
//                if (CT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    CT_RATIO_Text_Input_Layout.setError("This Field Is Required");
//                if (PT_RATIO_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    PT_RATIO_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_I_LOAD__PHASE_A_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_I_LOAD__PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_I_LOAD__PH_C_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_V_LOAD_PH_A_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_V_LOAD_PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    REAL_VALUE_V_LOAD_PH_C_Text_Input_Layout.setError("This Field Is Required");
//                if (I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    I_ONLOADED_BY_CLAMP_PH_A_Text_Input_Layout.setError("This Field Is Required");
//                if (I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    I_ONLOADED_BY_CLAMP_PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    I_ONLOADED_BY_CLAMP_PH_C_Text_Input_Layout.setError("This Field Is Required");
//                if (V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    V_ONLOAD_BY_V_M_V_LOAD_PH_A_Text_Input_Layout.setError("This Field Is Required");
//                if (V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    V_ONLOAD_BY_V_M_V_LOAD_PH_B_Text_Input_Layout.setError("This Field Is Required");
//                if (V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.getEditText().getText().toString().trim().isEmpty())
//                    V_ONLOAD_BY_V_M_V_LOAD_PH_C_Text_Input_Layout.setError("This Field Is Required");
                Toast.makeText(this, "Please Fill The Required Fields", Toast.LENGTH_SHORT).show();
            }
        }

        catch (Exception e) {
            Toast.makeText(this, "Not Saved :(", Toast.LENGTH_SHORT).show();
        }

    }

    public void SaveTheFile() {
        createPDF();
    }
}