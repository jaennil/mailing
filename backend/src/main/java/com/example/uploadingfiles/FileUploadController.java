package com.example.uploadingfiles;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.MvcUriComponentsBuilder;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.example.uploadingfiles.storage.StorageFileNotFoundException;
import com.example.uploadingfiles.storage.StorageService;

import javax.mail.*;
import javax.mail.internet.*;

@Controller
public class FileUploadController {

	private final StorageService storageService;

	@Autowired
	public FileUploadController(StorageService storageService) {
		this.storageService = storageService;
	}

	@GetMapping("/")
	public String listUploadedFiles(Model model) throws IOException {

		model.addAttribute("files", storageService.loadAll().map(
				path -> MvcUriComponentsBuilder.fromMethodName(FileUploadController.class,
						"serveFile", path.getFileName().toString()).build().toUri().toString())
				.collect(Collectors.toList()));

		return "uploadForm";
	}

	@GetMapping("/files/{filename:.+}")
	@ResponseBody
	public ResponseEntity<Resource> serveFile(@PathVariable String filename) {

		Resource file = storageService.loadAsResource(filename);
		return ResponseEntity.ok().header(HttpHeaders.CONTENT_DISPOSITION,
				"attachment; filename=\"" + file.getFilename() + "\"").body(file);
	}

	@PostMapping("/")
	public String handleFileUpload(@RequestParam("file") MultipartFile file, @RequestParam(required = false, name="emailText") String emailText, @RequestParam(required = false, name="emailSubject") String emailSubject) throws MessagingException, IOException {
		System.out.println(emailText);
		System.out.println("get file");
		storageService.store(file);

		// send email
		Properties prop = new Properties();
		prop.put("mail.smtp.auth", true);
		prop.put("mail.smtp.ssl.enable", "true");
		prop.put("mail.smtp.host", "smtp.mail.ru");
		prop.put("mail.smtp.port", "465");
		prop.put("mail.smtp.ssl.trust", "smtp.mail.ru");
		prop.put("mail.smtp.ssl.protocols", "TLSv1.2");
		Session session = Session.getInstance(prop, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication("n.e.dubrovskih@mail.ru", "mqDyAFMWLrYY1eb3HfXW");
			}
		});
		Message message = new MimeMessage(session);
		boolean flagEmail = false;
		boolean flagChiefName = false;
		String Email = "";
		String ChiefName = "";
		String datasetPath = ".\\datasets\\testdataset.xlsx";
		FileInputStream dataset = new FileInputStream(new File(datasetPath));

		Workbook workbook = new XSSFWorkbook(dataset);

		Sheet sheet = workbook.getSheetAt(0);

		Map<Integer, String> data = new HashMap<Integer, String>();
		int i = 0;

		for (Row row : sheet)
		{
			for (Cell cell : row)
			{
				switch (cell.getCellType())
				{
					case STRING:

						Pattern patternEmail = Pattern.compile("Email:[\\w+-.]+@\\w+.\\w+.\\w+");
						Matcher matcherEmail = patternEmail.matcher(cell.getRichStringCellValue().getString());

						if (matcherEmail.find())
						{
							Email = cell.getRichStringCellValue().getString().substring(6);
							flagEmail = true;
						}

						Pattern patternChiefName = Pattern.compile("ChiefName\\:[а-яА-яё]+\\s{1}[а-яА-яё]+\\s{1}[а-яА-яё]+");
						Matcher matcherChiefName = patternChiefName.matcher(cell.getRichStringCellValue().getString());

						if (matcherChiefName.find())
						{
							ChiefName = cell.getRichStringCellValue().getString().substring(matcherChiefName.start(), matcherChiefName.end()).substring(10);
							flagChiefName = true;
						}
						if (flagEmail && flagChiefName)
							data.put(i,Email + ":" + ChiefName);
						break;
					default:;
				}
			}
			i++;
			flagEmail = false;
			flagChiefName = false;
		}

		ArrayList<String> values = new ArrayList<>(data.values());
		String[] EmailAndChief;

		message.setFrom(new InternetAddress("n.e.dubrovskih@mail.ru"));

		for (String Element:values)
		{
			Element = Element.replaceAll("\n","");
			EmailAndChief = Element.split(":");

			message.setRecipients(
					Message.RecipientType.TO, InternetAddress.parse(EmailAndChief[0]));
			message.setSubject(emailSubject);

			MimeBodyPart mimeBodyPart = new MimeBodyPart();
			mimeBodyPart.setContent("Здравствуйте, " + EmailAndChief[1] + ", " + emailText, "text/html; charset=utf-8");

			Multipart multipart = new MimeMultipart();
			multipart.addBodyPart(mimeBodyPart);

			message.setContent(multipart);
			MimeBodyPart attachmentBodyPart = new MimeBodyPart();
			attachmentBodyPart.attachFile(new File(".\\upload-dir\\" + file.getOriginalFilename()));
			multipart.addBodyPart(attachmentBodyPart);
			Transport.send(message);
			System.out.println("message sent");
		}

		File f = new File(".\\upload-dir\\" + file.getOriginalFilename());
		if (f.delete()) {
			System.out.println(f.getName() + " deleted");
		} else {
			System.out.println("failed");
		}

		return "redirect:/";
	}

	@ExceptionHandler(StorageFileNotFoundException.class)
	public ResponseEntity<?> handleStorageFileNotFound(StorageFileNotFoundException exc) {
		return ResponseEntity.notFound().build();
	}

}
