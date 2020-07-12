package com.guilanzamineng;

import android.os.Bundle;
import android.os.Environment;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ArrayAdapter;
import android.widget.ListView;
import android.widget.TextView;

import androidx.fragment.app.Fragment;

import com.baoyz.widget.PullRefreshLayout;

import java.io.File;

public class PDF_Fragment extends Fragment {

    ListView list;
    PullRefreshLayout Refresh;
    TextView listv_status;

    @Override
    public View onCreateView(LayoutInflater inflater, ViewGroup container, Bundle savedInstanceState) {
        final View view = inflater.inflate(R.layout.list_fragment, container, false);

        list = (ListView) view.findViewById(R.id.listv);

        listv_status = view.findViewById(R.id.status);
        try{
            list.setVisibility(View.VISIBLE);
            listv_status.setVisibility(View.GONE);
            String path = Environment.getExternalStorageDirectory().toString()+"/Electricity/PDF/";
            File f = new File(path);//converted string object to file
            String[] values = f.list();//getting the list of files in string array
            //now presenting the data into screen
            assert values != null;
            if (values.length != 0) {
                ArrayAdapter<String> adapter = new ArrayAdapter<String>(getActivity().getApplicationContext(),android.R.layout.simple_spinner_dropdown_item, values);
                list.setAdapter(adapter);
            }
            else{
                list.setVisibility(View.GONE);
                listv_status.setVisibility(View.VISIBLE);
                listv_status.setText("There is no File To Show");
            }
        }
        catch (Exception e){
            list.setVisibility(View.GONE);
            listv_status.setVisibility(View.VISIBLE);
            listv_status.setText("Permissions Not Granted");
        }

        Refresh = view.findViewById(R.id.swipeRefreshLayout);

        Refresh.setOnRefreshListener(new PullRefreshLayout.OnRefreshListener() {
            @Override
            public void onRefresh() {
                String path = Environment.getExternalStorageDirectory().toString()+"/Electricity/PDF/";
                File f = new File(path);
                String[] values_refresh = f.list();
                ArrayAdapter<String> adapter = new ArrayAdapter<String>(getActivity(),android.R.layout.simple_spinner_dropdown_item, values_refresh);
                list.setAdapter(adapter);
                Refresh.setRefreshing(false);
            }
        });

        return view;
    }
}