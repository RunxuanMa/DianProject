package com.Controller;

import com.service.ExcelService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import javax.websocket.server.PathParam;
import java.io.*;
import java.net.URLEncoder;

@RestController
public class ExcelControl {

    String baseFilePath="D:\\fileSql\\";

    @Autowired
    private ExcelService excelService;

    @RequestMapping(value = "/Excel",method = RequestMethod.POST)
    public ModelAndView postExcel(MultipartFile file){

        String originalFilename = file.getOriginalFilename();

        String filePath="d:\\fileWork\\"+originalFilename;

        try {
            file.transferTo(new File(filePath));
            excelService.ReadDataFromExcelAndWriteInToDB(filePath,originalFilename);
        } catch (Exception e) {
            System.out.println(e.getMessage());
            ModelAndView mv = new ModelAndView();
            mv.addObject("errorMsg",e);
            mv.setViewName("/error");
            return mv;
        }

        return new ModelAndView("/success");

    }

    @GetMapping("/Excel")
    public ModelAndView getExcel(@RequestParam(value = "filename") String filename, HttpServletResponse response) throws IOException {
        File toFile = new File(baseFilePath+filename);
        String path=baseFilePath+filename;
       // System.out.println("filename="+filename);
        try {
            OutputStream os = new FileOutputStream(toFile);
            excelService.createExcel(filename).write(os);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
            ModelAndView mv = new ModelAndView();
            mv.addObject("errorMsg",e);
            mv.setViewName("/error");
            return mv;
        }
// 读到流中
     //   System.out.println("path="+path);
        InputStream inputStream = new FileInputStream(path);// 文件的存放路径
        response.reset();
        response.setContentType("application/octet-stream");
        String filename01 = new File(path).getName();
        response.addHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(filename01, "UTF-8"));
        ServletOutputStream outputStream = response.getOutputStream();
        byte[] b = new byte[1024];
        int len;
//从输入流中读取一定数量的字节，并将其存储在缓冲区字节数组中，读到末尾返回-1
        while ((len = inputStream.read(b)) > 0) {
            outputStream.write(b, 0, len);
        }
        inputStream.close();
        outputStream.close();

        return new ModelAndView("/success");

    }



}
