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

    Recruiter(String name, String email) {
        this.name = name;
        this.email = email;
    }
}

public class EmailSender {
    private static final int START_INDEX = 31;   // Change to desired start row
    private static final int END_INDEX = 32;   // Change to desired end row

    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\kundu\\Dropbox\\PC\\Desktop\\cold-mail.xlsx";
        List<Recruiter> recruiters = readEmailsFromExcel(excelFilePath, START_INDEX, END_INDEX);

        for (Recruiter recruiter : recruiters) {
            sendEmail(recruiter);
            try {
                Thread.sleep(5000); // 5-second delay to avoid spam detection
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

                Cell nameCell = row.getCell(2);  // Column B (Name)
                Cell emailCell = row.getCell(3); // Column C (Email)

                String name = getCellValueAsString(nameCell);
                String email = getCellValueAsString(emailCell).trim(); // Fix for email retrieval

                if (name != null && !email.isEmpty()) {
                    recruiters.add(new Recruiter(name, email));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return recruiters;
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue()); // Convert to long to remove decimals
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> null;
        };
    }

    public static void sendEmail(Recruiter recruiter) {

        Dotenv dotenv = Dotenv.configure()
                .directory("src/main/resources")
                .load();

        final String username = dotenv.get("EMAIL_USER");// Your email
        final String password = dotenv.get("EMAIL_PASSWORD"); // Use environment variable for security

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
            message.setSubject("Full Stack Java Web Developer");

            String emailContent = "Dear " + recruiter.name + ",\n\n" +
                    "I hope this message finds you well. My name is Vineet Kundu, and I am a final-year student at the World College of Technology and Management, FarukhNagar, Gurgaon.\n" +
                    "\n" +
                    "I am writing to express my keen interest in a job opportunity in the field of Software Development. With a strong passion for software development, particularly in React and Java, I am eager to contribute to innovative projects and gain practical experience in a professional setting.\n" +
                    "\n" +
                    "I have 9 months of experience as a React and Java Developer at Qubitnets Technologies, where I honed my skills in building scalable applications and working collaboratively within a team. Additionally, I am proficient in databases, Java, Spring, and Spring Boot, and I am available to join immediately.\n" +
                    "\n" +
                    "You can find more about my work and contributions through the following profiles:\n" +
                    "\n" +
                    "GitHub: https://github.com/KunduVineet\n" +
                    "\n" +
                    "LinkedIn: https://www.linkedin.com/in/vineet-kundu-b83407218/\n" +
                    "\n" +
                    "Portfolio: https://vk-portfolio-cd9d.vercel.app/ \n" +
                    "\n" +
                    "Resume : https://drive.google.com/file/d/1vCOIKf_a0sQZm1YApkNMAcTbp7xvYOm3/view?usp=drivesdk \n" +
                    "\n" +
                    "\n" +
                    "I would greatly appreciate the opportunity to discuss potential job openings or the application process further. Please find my resume attached for your reference.\n" +
                    "\n" +
                    "Thank you for considering my application. I look forward to the possibility of contributing to your team and its projects.\n" +
                    "\n" +
                    "Best regards,\n" +
                    "Vineet Kundu\n" +
                    "kunduvineet6@gmail.com\n" +
                    "+91 8882924671";

            message.setText(emailContent);
            Transport.send(message);

            System.out.println("✅ Email sent to: " + recruiter.name + " <" + recruiter.email + ">");
        } catch (MessagingException e) {
            System.out.println("❌ Failed to send email to: " + recruiter.name);
            e.printStackTrace();
        }
    }
}
