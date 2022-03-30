package com.example.demo;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

@RequestMapping("test")
@RestController
public class TestController {

    @GetMapping("1")
    public ModelAndView test1() {
        Map<String, Object> map = new HashMap<>();
        map.put("name", "张三");
        map.put("age", "18");
        return new ModelAndView(new JxlsView("templates/template1.xlsx", "output1"), map);
    }

    @GetMapping(value = "2")
    public void test2(HttpServletRequest request, HttpServletResponse response) throws IOException {

        List<User> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            User user = new User();
            user.setId(String.valueOf(i));
            user.setName("name" + i);
            userList.add(user);
        }

        InputStream is = new ClassPathResource("templates/template2.xlsx").getInputStream();
        OutputStream os = null;
        try {
            os = response.getOutputStream();
            response.addHeader("Content-Disposition", "attachment;filename=" + "output2.xlsx");
            response.setContentType("application/octet-stream");
            Context context = new Context();
            context.putVar("userList", userList);
            JxlsHelper.getInstance().processTemplate(is, os, context);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            os.flush();
            os.close();
        }
    }

    @GetMapping(value = "3")
    public void test3(HttpServletRequest request, HttpServletResponse response) throws IOException {

        List<User> userList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            User user = new User();
            user.setId(String.valueOf(i));
            user.setName("name" + i);
            // if (i % 2 == 0) {
            //     user.setType("偶数");
            // } else {
            //     user.setType("奇数");
            // }
            userList.add(user);
        }

        InputStream is = new ClassPathResource("templates/template3.xlsx").getInputStream();
        OutputStream os = null;
        try {
            os = response.getOutputStream();
            response.addHeader("Content-Disposition", "attachment;filename=" + "output3.xlsx");
            response.setContentType("application/octet-stream");
            Context context = new Context();
            context.putVar("userList", userList);
            JxlsHelper.getInstance().processTemplate(is, os, context);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            os.flush();
            os.close();
        }
    }

    @GetMapping(value = "4")
    public void test4(HttpServletRequest request, HttpServletResponse response) throws IOException {

        List<User> userList = new ArrayList<>();
        userList.add(new User("1", "yuji", 60D));
        userList.add(new User("1", "zhangchunhua", 120D));
        userList.add(new User("1", "caifuren", 120D));
        userList.add(new User("2", "zhouyu", 100D));
        userList.add(new User("2", "sunce", 110D));
        userList.add(new User("3", "machao", 130D));
        userList.add(new User("3", "zuoci", 150D));

        InputStream is = new ClassPathResource("templates/template4.xlsx").getInputStream();
        OutputStream os = null;
        try {
            os = response.getOutputStream();
            response.addHeader("Content-Disposition", "attachment;filename=" + "output4.xlsx");
            response.setContentType("application/octet-stream");
            Context context = new Context();
            context.putVar("userList", userList);
            JxlsHelper.getInstance().processTemplate(is, os, context);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            os.flush();
            os.close();
        }
    }

    // 单元格合并
    // public WritableCellValue mergeCell(String value , Integer mergerRows) {
    //     return new MergeCellValue(value , mergerRows);
    // }

}
