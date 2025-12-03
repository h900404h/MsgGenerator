package demo.msggenerator;

import com.opencsv.CSVReader;

import ch.astorm.jotlmsg.OutlookMessage;
import ch.astorm.jotlmsg.OutlookMessageAttachment;
import ch.astorm.jotlmsg.OutlookMessageRecipient;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class MsgGenerator {

    // CSV 路徑：mail_list.csv（與執行 jar 同層）
    private static Path getCsvPath() {
        return Paths.get("mail_list.csv").toAbsolutePath();
    }

    // 附件資料夾：table（公開版不含任何公司名稱）
    private static Path getAttachmentDir() {
        return Paths.get("table").toAbsolutePath();
    }

    // 輸出資料夾：output_msg/YYYY/MM
    private static Path getOutDir(LocalDate date) throws IOException {
        String year  = date.format(DateTimeFormatter.ofPattern("yyyy"));
        String month = date.format(DateTimeFormatter.ofPattern("MM"));
        Path dir = Paths.get("output_msg", year, month).toAbsolutePath();
        Files.createDirectories(dir);
        return dir;
    }

    // 簽名（已匿名化）
    private static final String SIGNATURE_TEXT = String.join("\r\n",
        "============================================",
        "Example Corporation  https://www.example.com",
        "Taro Yamada（タロウ ヤマダ）  E-Mail：example@example.com",
        "〒000-0000　東京都サンプル区サンプル町 0-0-0 サンプルビル 1F",
        "TEL：00-0000-0000　 FAX：00-0000-0000",
        "============================================"
    );

    public static void main(String[] args) throws Exception {

        // 東京時區日期與月份
        ZoneId tz = ZoneId.of("Asia/Tokyo");
        LocalDate today = LocalDate.now(tz);
        int month = today.getMonthValue();

        String bodyText = buildBodyText(month);

        Path csv = getCsvPath();
        Path attachDir = getAttachmentDir();

        if (!Files.exists(csv)) throw new FileNotFoundException("CSV not found: " + csv);
        if (!Files.isDirectory(attachDir)) throw new FileNotFoundException("Attachment folder not found: " + attachDir);

        try (CSVReader reader = new CSVReader(
                new InputStreamReader(Files.newInputStream(csv), StandardCharsets.UTF_8))) {

            String[] first = reader.readNext();
            if (first == null) {
                throw new IllegalStateException("CSV is empty: " + csv);
            }

            boolean isHeader = String.join(",", first).toLowerCase().contains("to")
                            || String.join(",", first).toLowerCase().contains("cc");

            boolean hasData = false;

            if (!isHeader) {
                processRow_B(first, month, today, attachDir, bodyText);
                hasData = true;
            }

            String[] row;
            while ((row = reader.readNext()) != null) {
                hasData = true;
                processRow_B(row, month, today, attachDir, bodyText);
            }

            if (!hasData) {
                throw new IllegalStateException("No data rows found in CSV: " + csv);
            }
        }
    }

    private static List<Path> findAttachmentsByName(Path dir, String keyword) throws IOException {
        if (keyword == null || keyword.isBlank()) return List.of();

        try (var stream = Files.list(dir)) {
            return stream
                .filter(Files::isRegularFile)
                .filter(p -> p.getFileName().toString().contains(keyword))
                .sorted()
                .toList();
        }
    }

    /**
     * CSV 欄位：to,cc,filename_keyword
     * 一行一封信，每位收件人一封。
     */
    private static void processRow_B(String[] row, int month, LocalDate today, Path attachDir,
                                     String bodyText) throws IOException {

        String to = safe(row, 0);
        String ccForRow = safe(row, 1);
        String suffixFromCsv = safe(row, 2);

        if (to.isBlank()) {
            throw new IllegalArgumentException("CSV error: 'to' column cannot be blank.");
        }

        String suffix = suffixFromCsv;
        List<Path> attachments = findAttachmentsByName(attachDir, suffix);

        if (attachments.isEmpty()) {
            throw new FileNotFoundException(
                "No matching attachments for keyword: \"" + suffix + "\" in: " + attachDir
            );
        }

        for (String addr : splitRecipients(to)) {
            String email = addr.trim();
            if (email.isBlank()) continue;

            String subjectPersonal = month + "月 支払明細";

            String fileSafeSuffix = suffixFromCsv.replaceAll("[\\\\/:*?\"<>|\\s]+", "_");

            String filename = DateTimeFormatter.ofPattern("yyyyMMdd").format(today)
                    + "_" + fileSafeSuffix + ".msg";
            Path savePath = getOutDir(today).resolve(filename).toAbsolutePath();

            OutlookMessage msg = new OutlookMessage();
            msg.setSubject(subjectPersonal);
            msg.setPlainTextBody(bodyText);

            msg.addRecipient(OutlookMessageRecipient.Type.TO, email);

            for (String ccAddr : splitRecipients(ccForRow)) {
                if (!ccAddr.isBlank()) {
                    msg.addRecipient(OutlookMessageRecipient.Type.CC, ccAddr.trim());
                }
            }

            for (Path p : attachments) {
                OutlookMessageAttachment att = new OutlookMessageAttachment(
                    p.getFileName().toString(),
                    "application/vnd.ms-excel.sheet.macroEnabled.12",
                    Files.newInputStream(p)
                );
                msg.addAttachment(att);
            }

            msg.writeTo(savePath.toFile());
            System.out.println("Saved: " + savePath + "  → TO=" + email + "  CC=" + ccForRow);
        }
    }

    private static String buildBodyText(int month) {
        return String.join("\r\n\r\n",
            "お疲れさまです。",
            month + "月の支払明細を送付します。",
            "ご確認よろしくお願いいたします。"
        ) + "\r\n\r\n" + SIGNATURE_TEXT + "\r\n";
    }

    private static String safe(String[] arr, int idx) {
        return (idx < arr.length) ? arr[idx].trim() : "";
    }

    private static String[] splitRecipients(String s) {
        if (s == null) return new String[0];
        return s.replace('，',';')
                .replace('；',';')
                .replace('、',';')
                .replace(',', ';')
                .split(";");
    }
}
