package com.caputo.testExcel.controller;

import com.caputo.testExcel.entities.Client;
import com.caputo.testExcel.exporter.ExcelFileExporter;
import org.apache.commons.compress.utils.IOUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
@RequestMapping(value = "/excels")
public class DownloadExcel {

    public String index(){
        return "index";
    }

    @GetMapping("/download")
    public void downloadExcelFile(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=clients.xlsx");

        ByteArrayInputStream inputStream = ExcelFileExporter.exportCustomerListToExcelFile(createTestData());

        IOUtils.copy(inputStream, response.getOutputStream());

    }

    private List<Client> createTestData(){
        List<Client> clients = new ArrayList<>();
        clients.add(new Client(1, 100, "Jon"));
        clients.add(new Client(2, 100, "Michel"));
        clients.add(new Client(3, 100, "Rafael"));
        return clients;
    }

}
