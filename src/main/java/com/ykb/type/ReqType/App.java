package com.ykb.type.ReqType;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Base64.Encoder;
import java.util.Date;

import javax.annotation.Resource;

import org.apache.axis.encoding.Base64;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.http.converter.ByteArrayHttpMessageConverter;
import org.springframework.http.converter.FormHttpMessageConverter;
import org.springframework.http.converter.json.MappingJackson2HttpMessageConverter;
import org.springframework.web.client.RestClientException;
import org.springframework.web.client.RestTemplate;

import com.ykb.external.services.externalservices.deployment.ConnectionProvider;
import com.ykb.external.services.externalservices.deployment.ConnectionProviderImplService;
import com.ykb.external.services.externalservices.deployment.ServiceRequest4GetConnectionProvider;
import com.ykb.external.services.externalservices.deployment.ServiceResponse4GetConnectionProvider;

import https.tempuri.PasswordResponse;

public class App {

	private String userName = "xxxx";
	private String password = encPass();
	private static Logger logger=LoggerFactory.getLogger(App.class);
	private static void mailSend(File f)

	{
		try {
			URI url = new URI("rest_servis_url");
			
			FileInputStream inputFile = new FileInputStream(f);
			byte fileData[] = new byte[(int) f.length()];
			inputFile.read(fileData);
			String encodedFile = Base64.encode(fileData);
			// String decodeFName=URLDecoder.decode(f.getName(), "UTF-8");
			// String fileName = URLDecoder.decode(f.getPath(),
			// String.valueOf(StandardCharsets.UTF_8));
			Long longSize = new Long(f.length());
			int size = longSize.intValue();
			RestTemplate restTemplate = new RestTemplate();
			MailModel newMail = new MailModel();
			newMail.setAnalyst("Melisa Polat");
			newMail.setApplication("KKBApp");
			newMail.setCcList("melisapolat.03@hotmail.com");
			newMail.setContent("TWVyaGFiYWxhciAKRGVuZW1lIG1haWxpZGlyLgrEsHlpIMOHYWzEscWfbWFsYXIK");
			newMail.setFromAddress("melisapolt@gmail.com");
			newMail.setFromName("Melisa Polat");
			newMail.setReplyToAddress("null");
			newMail.setReplyToName("null");
			newMail.setSubject("Kkb Müşteri Çalışma Bildirim Formu");
			newMail.setToList("melisapolt@gmail.com");
			newMail.setUser("DenemeMerkezi");

			Attachment attmnt2 = new Attachment();
			attmnt2.setContent(encodedFile);
			attmnt2.setName(f.getName());
			attmnt2.setSize(size);
			Attachment[] attmntList = new Attachment[] { attmnt2 };
			newMail.setAttachments(attmntList);

			HttpEntity<MailModel> entity = new HttpEntity<MailModel>(newMail);
			// System.out.println(entity);
			
			ResponseEntity<ServiceResp> result = restTemplate.postForEntity(url, entity, ServiceResp.class);
			logger.info("Message :" + result.getBody().getSendStatus());
		//	System.out.println("Message : " + result.getBody().getSendStatus());

		} catch (Exception e) {
			logger.error(e.toString());
			//System.out.println("" + e.toString());
		}
	}
	
	public static void main(String[] args) throws URISyntaxException  {
		new App().export();
	}

	public void export() {
		String jdbcURL = "jdbc:jtds:sqlserver://server_name:port/db_name;instance=instance_name";

		try (Connection connection = DriverManager.getConnection(jdbcURL, userName, password)) {

			String reqSql = "Select a.request_id from z_usm_request_item_form a left outer join zKKB_Mail b  on a.request_id=b.request_id left outer join usm_request ur on a.request_id=ur.request_id where a.form_elem_name like 'kkb_talepsorumlusu' and b.request_id is null and ur.status=2" ;
			Statement statement1 = connection.createStatement();
			ResultSet reqResult = statement1.executeQuery(reqSql);
			ResultSetMetaData rsmd = reqResult.getMetaData();
			DateFormat sdf = new SimpleDateFormat("dd.MM.YYYY");
			Date tarih = new Date();
			

			String countSql="Select count(*) count from z_usm_request_item_form a left outer join zKKB_Mail b  on a.request_id=b.request_id left outer join usm_request ur on a.request_id=ur.request_id where a.form_elem_name like 'kkb_talepsorumlusu' and b.request_id is null and ur.status=2";
			Statement statementCount=connection.createStatement();
			ResultSet countResult=statementCount.executeQuery(countSql);
			countResult.next();
			int rowcount=countResult.getInt("count");
			statementCount.close();
			int i=1;
			while (reqResult.next()) {
				int reqId = reqResult.getInt("request_id");
				String directoryName="xxxx";
				String fileSeparator=System.getProperty("file.separator");
				String relativePath=directoryName+fileSeparator+"Müşteri Çalışma Bildirim Formu.xlsx";
				
				File file1=new File(relativePath);
				FileInputStream inputstream = new FileInputStream(file1);

					// ü,ş,ç,ı girmediğim zaman düzgün okuyor dosya ismini!
				
					String excelFilePath =directoryName+fileSeparator+ ("Musteri Calisma Bildirim Formu ") + sdf.format(tarih) + " (00" + i
							+ ").xlsx";

					String sql = "SELECT * FROM z_usm_request_item_form where request_id=" + reqId
							+ " and (form_elem_name like 'kkb%' or form_elem_name like '%desc')";

					Statement statement = connection.createStatement();
					ResultSet result = statement.executeQuery(sql);
					XSSFWorkbook workbook = new XSSFWorkbook(inputstream);
					XSSFSheet sheet = workbook.getSheet("Müşteri Bildirim");
					writeDataLines(result, workbook, sheet);
					FileOutputStream outputStream = new FileOutputStream(excelFilePath);
					workbook.write(outputStream);
					i++;
					File f = new File(excelFilePath);
					mailSend(f);
					workbook.close();
					inputstream.close();
					String insertSql="Insert into zKKB_Mail (request_id,islem_tarihi) values ("+reqId+",GetDate())";
					Statement statementInsert = connection.createStatement();
					statementInsert.executeUpdate(insertSql);
					statementInsert.close();
					statement.close();
			}
			
			statement1.close();

		} catch (SQLException e) {
			logger.error("Database error "+e.toString());
			//System.out.println("Database error:");
			//e.printStackTrace();
		} catch (IOException e) {
			logger.error("ERROR: "+e);
		//	System.out.println(e);
			
		}
	}

	private void writeDataLines(ResultSet result, XSSFWorkbook workbook, XSSFSheet sheet) throws SQLException {
		int rowCount = 1;

		while (result.next()) {
			String elemanName = result.getString("form_elem_name");
			String elemanValue = result.getString("form_elem_value");

			int columnCount = 0;
			Row row = sheet.getRow(rowCount + 1);

			Cell cell;

			if (elemanName.equals("kkb_bildirimtipi")) {
				row = sheet.getRow(rowCount + 4);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_calismafirmalar")) {
				row = sheet.getRow(rowCount + 6);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_calismasistem")) {
				row = sheet.getRow(rowCount + 11);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_calismasuresi")) {
				row = sheet.getRow(rowCount + 10);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_calismatarihi")) {
				row = sheet.getRow(rowCount + 9);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_calismayapacaklar")) {
				row = sheet.getRow(rowCount + 7);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_destekvar")) {
				row = sheet.getRow(rowCount + 12);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_digernot")) {
				row = sheet.getRow(rowCount + 14);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_firma")) {
				row = sheet.getRow(rowCount + 1);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_plaka")) {
				row = sheet.getRow(rowCount + 8);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_talepkategorisi")) {
				row = sheet.getRow(rowCount + 5);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_talepsorumlusu")) {
				row = sheet.getRow(rowCount + 2);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("kkb_tel")) {
				row = sheet.getRow(rowCount + 3);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			} else if (elemanName.equals("req_desc")) {
				row =sheet.getRow(rowCount + 13);
				cell = row.getCell(columnCount++);
				// cell.setCellValue(elemanName);
			}

			cell = row.getCell(columnCount);
			if (elemanValue.contains("|")) {
				String[] elemanVal = elemanValue.split("\\|");
				String elemanVal2 = elemanVal[0];
				// System.out.println(elemanVal2);
				cell.setCellValue(elemanVal2);

			} else {
				cell.setCellValue(elemanValue);
			}
		}
	}

	public String encPass() {
		String adress = "server_name";
		ServiceRequest4GetConnectionProvider request = new ServiceRequest4GetConnectionProvider();
		request.setUserName(userName);
		request.setAddress(adress);

		ConnectionProvider cp = new ConnectionProviderImplService().getConnectionProviderImplPort();
		ServiceResponse4GetConnectionProvider response = cp.getConnectionProvider(request);
		PasswordResponse passwordResponse = response.getResp();

		String encPassword = passwordResponse.getContent();
		encPassword.getBytes();
		byte[] byteArray = Base64.decode(encPassword);

		String password = new String(byteArray);
		return password;

	}
}
