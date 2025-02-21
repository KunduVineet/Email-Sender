package org.example;

import io.github.cdimascio.dotenv.Dotenv;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.mail.*;
import javax.mail.internet.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

class Recruiter {
    String name;
    String email;
    String companyName;
    String title;

    Recruiter(String name, String email, String companyName, String title) {
        this.name = name;
        this.email = email;
        this.companyName = companyName;
        this.title = title;
    }
}

public class EmailSender {
    private static final int START_INDEX = 101;   // Change to desired start row
    private static final int END_INDEX = 300;   // Change to desired end row

    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\kundu\\Dropbox\\PC\\Desktop\\cold-mail.xlsx";
        List<Recruiter> recruiters = readEmailsFromExcel(excelFilePath, START_INDEX, END_INDEX);

        for (Recruiter recruiter : recruiters) {
            sendEmail(recruiter);
            try {
                Thread.sleep(8000); // 8-second delay to avoid spam detection
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }

    public static List<Recruiter> readEmailsFromExcel(String filePath, int start, int end) {
        List<Recruiter> recruiters = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int rowCount = sheet.getPhysicalNumberOfRows();

            for (int i = start; i <= end && i < rowCount; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell nameCell = row.getCell(2);  // Column C (Name)
                Cell emailCell = row.getCell(3); // Column D (Email)
                Cell titleCell = row.getCell(4); // Column E (Title)
                Cell companyCell = row.getCell(5); // Column F (Company)

                String name = getCellValueAsString(nameCell);
                String email = getCellValueAsString(emailCell).trim();
                String title = getCellValueAsString(titleCell);
                String companyName = getCellValueAsString(companyCell).trim();

                if (name != null && !email.isEmpty()) {
                    recruiters.add(new Recruiter(name, email, companyName, title));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return recruiters;
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue()); // Convert to long to remove decimals
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    public static void sendEmail(Recruiter recruiter) {
        Dotenv dotenv = Dotenv.configure()
                .directory("src/main/resources")
                .load();

        final String username = dotenv.get("EMAIL_USER");
        final String password = dotenv.get("EMAIL_PASSWORD");

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        Session session = Session.getInstance(props, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, password);
            }
        });

        try {
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(username));
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(recruiter.email));
            message.setSubject("Full Stack Java Web Developer - " + recruiter.companyName);

            String emailContent = "Dear " + recruiter.name + ",\n\n" +
                    "I hope this message finds you well. MySelf Vineet Kundu from MDU, and I am reaching out to express my keen interest in a job opportunity at " + recruiter.companyName + ".\n\n" +
                    "I noticed that you are a " + recruiter.title + " at " + recruiter.companyName + ", and I would love the opportunity to discuss how my Skills and Experience align with your team’s needs.\n\n" +
                    "With 9 months of Experience as a React and Java Developer at Qubitnets Technologies, I have honed my skills in developing Scalable Applications. Additionally, I am proficient in Java, Spring Boot, and databases, and I am available for immediate joining.\n\n" +
                    "Here are some links to my work:\n" +
                    "GitHub: https://github.com/KunduVineet\n\n" +
                    "LinkedIn: https://www.linkedin.com/in/vineet-kundu-b83407218/\n\n" +
                    "Portfolio: https://vk-portfolio-cd9d.vercel.app/\n\n" +
                    "Resume: https://drive.google.com/file/d/10ryvX2HN05xkU3Is9Odz2hGpY2yvGsO-/view?usp=drivesdk\n\n" +
                    "I would love to connect and discuss potential opportunities at " + recruiter.companyName + ". Please let me know a convenient time to chat.\n\n" +
                    "Best regards,\n" +
                    "Vineet Kundu\n" +
                    "kunduvineet6@gmail.com\n" +
                    "+91 8882924671";

            message.setText(emailContent);
            Transport.send(message);

            System.out.println("✅ Email sent to: " + recruiter.name + " <" + recruiter.email + ">");
        } catch (MessagingException e) {
            System.out.println("❌ Failed to send email to: " + recruiter.name + " <" + recruiter.email + ">");
            e.printStackTrace();
        }
    }
}
