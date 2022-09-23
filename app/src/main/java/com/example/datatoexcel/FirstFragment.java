package com.example.datatoexcel;

import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;

import androidx.annotation.NonNull;
import androidx.fragment.app.Fragment;
import androidx.navigation.fragment.NavHostFragment;

import com.example.datatoexcel.databinding.FragmentFirstBinding;
import com.example.datatoexcel.utils.ExcelUtil;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class FirstFragment extends Fragment {

    private FragmentFirstBinding binding;

    @Override
    public View onCreateView(LayoutInflater inflater, ViewGroup container,
                             Bundle savedInstanceState) {
        binding = FragmentFirstBinding.inflate(inflater, container, false);
        return binding.getRoot();
    }

    public void onViewCreated(@NonNull View view, Bundle savedInstanceState) {
        super.onViewCreated(view, savedInstanceState);

        binding.buttonFirst.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                NavHostFragment.findNavController(FirstFragment.this)
                        .navigate(R.id.action_FirstFragment_to_SecondFragment);
            }
        });

        binding.btnToExcel.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                exportExcel();
            }
        });
    }

    @Override
    public void onDestroyView() {
        super.onDestroyView();
        binding = null;
    }

    private void exportExcel() {
        // 表格标题
        String[] title = {"场地费支出", "水费支出", "电费支出", "员工报销支出", "客户收入", "公司授权收入", "月份"};
        // 表格文件名
        String fileName = System.currentTimeMillis() + "_" + "财务收支汇总表.xls";
        // 初始化表格
        ExcelUtil.initExcel(getContext(), fileName, "财务记录", title);

        List<List<String>> dataList = new ArrayList<>();
        // 第一行数据
        List<String> rowData = Arrays.asList("54000", "560", "1004", "20000", "1040000", "23000", "1");
        // 第二行数据
        List<String> rowData1 = Arrays.asList("54000", "580", "1084", "20200", "1240000", "23000", "2");
        // 第三行数据
        List<String> rowData2 = Arrays.asList("54000", "520", "1504", "20078", "1340000", "23000", "3");
        dataList.add(rowData);
        dataList.add(rowData1);
        dataList.add(rowData2);

        // 数据写入表格
        ExcelUtil.writeObjListToExcel(dataList, fileName, getContext());
    }

}