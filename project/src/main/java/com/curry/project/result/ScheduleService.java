package com.curry.project.result;


import jakarta.mail.internet.MimeMessage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.scheduling.annotation.Async;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.List;

@ConditionalOnProperty(value = "app.scheduling.enable", havingValue = "true", matchIfMissing = false)
@Component
public class ScheduleService {

    @Autowired
    private ResultService resultService;

    @Scheduled(cron = "0 0 9 1 * ?")
    public void monthlyExportTask() {
        System.out.println("--- 測試排程開始執行: " + LocalDateTime.now() + " ---");
        try {
            resultService.exportExcel();
            System.out.println("每月自動匯出排程已觸發成功。");
        } catch (Exception e) {
            System.err.println("自動匯出排程失敗: " + e.getMessage());
            e.printStackTrace();
        }
    }
}