# Java Email Sender From Excel Files

## 📌 Project Overview
This Java-based application reads recruiter details (name and email) from an Excel file and sends personalized cold emails automatically. It uses:
- **Apache POI** to read `.xlsx` files
- **JavaMail API** to send emails
- **dotenv-java** to securely manage environment variables

---

## 🚀 Features
✅ Reads recruiter names and emails from an Excel file  
✅ Sends automated cold emails using SMTP  
✅ Uses `.env` file for storing sensitive credentials  
✅ Delays email sending to avoid spam detection  
✅ Logs email delivery status in the console  

---

## 📂 Folder Structure
```
Email_Sending_Java/
│── src/
│   ├── main/
│   │   ├── java/org/example/
│   │   │   ├── EmailSender.java
│   │   │   ├── Recruiter.java
│   │   ├── resources/
│   │   │   ├── .env
│── .gitignore
│── pom.xml
│── README.md
```

---

## 🔧 Prerequisites
Make sure you have the following installed:
- **Java 8+** (JDK 11 or later recommended)
- **Apache Maven** (For dependency management)
- **Gmail SMTP App Password** (for secure email sending)
- **Excel file (.xlsx)** with recruiter details

---

## 🛠 Setup & Installation

### 1️⃣ Clone the Repository
```sh
git clone https://github.com/your-username/Email_Sending_Java.git
cd Email_Sending_Java
```

### 2️⃣ Create a `.env` File
Inside the `src/main/resources/` directory, create a file named `.env` and add the following:
```env
EMAIL_USER=your-email@gmail.com
EMAIL_PASSWORD=your-app-password
EXCEL_FILE_PATH=C:\Users\YourName\Desktop\cold-mail.xlsx
```
⚠ **Important:** Do not use your actual Gmail password. Use an [App Password](https://myaccount.google.com/apppasswords).

### 3️⃣ Add `.env` to `.gitignore`
To prevent credentials from being committed to Git:
```sh
echo ".env" >> .gitignore
git rm --cached .env
git commit -m "Ignore .env file"
```

### 4️⃣ Install Dependencies
Run Maven to install required dependencies:
```sh
mvn clean install
```

---

## 📌 How to Run
Run the application using IntelliJ IDEA or the command line:
```sh
mvn exec:java -Dexec.mainClass="org.example.EmailSender"
```

---

## 📜 Code Explanation
### **Recruiter.java**
A simple class to store recruiter details:
```java
class Recruiter {
    String name;
    String email;
    Recruiter(String name, String email) {
        this.name = name;
        this.email = email;
    }
}
```

### **EmailSender.java**
Handles reading Excel files and sending emails.

#### **Reading Excel File**
```java
public static List<Recruiter> readEmailsFromExcel(String filePath, int start, int end) {
    FileInputStream fis = new FileInputStream(new File(filePath));
    Workbook workbook = new XSSFWorkbook(fis);
    Sheet sheet = workbook.getSheetAt(0);
    ...
}
```

#### **Sending Email**
```java
Session session = Session.getInstance(props, new Authenticator() {
    protected PasswordAuthentication getPasswordAuthentication() {
        return new PasswordAuthentication(username, password);
    }
});
```

---

## 🛠 Troubleshooting
**1️⃣ Log4j2 Warning: "Could not find a logging implementation."**
- Add `log4j-core` to `pom.xml`:
  ```xml
  <dependency>
      <groupId>org.apache.logging.log4j</groupId>
      <artifactId>log4j-core</artifactId>
      <version>2.18.0</version>
  </dependency>
  ```

**2️⃣ `.env` File Not Found Error**
- Ensure `.env` is placed inside `src/main/resources/`.
- Run `mvn clean` and restart the application.

**3️⃣ Gmail SMTP Issues**
- Enable "Less secure apps" in Gmail or use an **App Password**.

---

## 📜 License
This project is licensed under the MIT License.

---

## 🤝 Contributing
Feel free to fork and improve this project! Pull requests are welcome.

---

## 📬 Contact
📧 Email: (mailto:kunduvineet6@gmail.com)  
🔗 GitHub: https://github.com/KunduVineet/
🔗 LinkedIn: https://www.linkedin.com/in/vineet-kundu-b83407218

