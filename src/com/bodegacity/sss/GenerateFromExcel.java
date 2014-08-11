package com.bodegacity.sss;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Hashtable;

import javax.print.Doc;
import javax.print.DocFlavor;
import javax.print.DocPrintJob;
import javax.print.PrintException;
import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.SimpleDoc;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Copies;
import javax.print.event.PrintJobAdapter;
import javax.print.event.PrintJobEvent;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import r3source.ApplicablePeriod;
import r3source.Contribution;
import r3source.Employee;
import r3source.Employer;
import r3source.Payment;
import r3source.R3File;
import r3source.Utility;

public class GenerateFromExcel {

	/**
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws Exception {
		File f = new File("");

		String path = f.getAbsolutePath();
		String sheetName = path.substring(path.lastIndexOf(File.separator) + 1);

		FileInputStream file = new FileInputStream(new File("../SSS.xlsx"));

		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(sheetName);

		Date applicableDate = sheet.getRow(2).getCell(2).getDateCellValue();
		String SBRNumber = sheet.getRow(4).getCell(2).getStringCellValue();

		double amount = sheet.getRow(5).getCell(2).getNumericCellValue();
		Date paymentDate = sheet.getRow(6).getCell(2).getDateCellValue();

		Calendar cal = new GregorianCalendar();
		cal.setTime(applicableDate);
		ApplicablePeriod applicablePeriod = new ApplicablePeriod(
																	cal.get(Calendar.MONTH) + 1,
																	cal.get(Calendar.YEAR));

		Employer employer = Utility.readER();
		Hashtable<String, Employee> employees = getEmployees(sheet, employer, applicablePeriod);

		SimpleDateFormat paymentSdf = new SimpleDateFormat("MMddYYYY");
		Payment payment = new Payment(
										applicablePeriod,
										SBRNumber,
										amount,
										paymentSdf.format(paymentDate));

		R3File r3File = new R3File(applicablePeriod, employer, employees, payment);

		SimpleDateFormat sdf = new SimpleDateFormat("MMddHHmm");
		String extension = sdf.format(new Date());
		String fileName = "R3" +
							r3File.getEmployer().getErSssNumber() +
							r3File.getAppPeriod().strImage() +
							"." +
							extension;
		File r3 = new File(fileName);
		r3File.createTextFile(r3);

		r3File.createEmployeeReport();
		r3File.createTransmittalReport();

		Utility.writeObject(r3File, new File("r3File.dat"));

		SimpleDateFormat folderSdf = new SimpleDateFormat("YYYY-MM");
		File folder = new File(folderSdf.format(applicableDate));
		if (!folder.exists()) {
			folder.mkdir();
		}
		File employeeFile = new File("EMPLOYEE_LIST");
		File newEmployeeFile = new File(folder, "EMPLOYEE_LIST");
		move(employeeFile, newEmployeeFile);
		
		File transmittalFile = new File("TRANSMITTAL_REPORT");
		File newTransmittalFile = new File(folder, "TRANSMITTAL_REPORT");
		move(transmittalFile, newTransmittalFile);
		
		Files.copy(r3.toPath(), new FileOutputStream(new File("../_usb/" + fileName)));
		File newR3File = new File(folder, fileName);
		move(r3, newR3File);

		boolean print = args.length > 0 ? Boolean.valueOf(args[0]) : false;
		if (print) {
			print(newTransmittalFile, newEmployeeFile);
		}
	}

	private static void move(File source, File dest) {
		if (dest.exists()) {
			dest.delete();
		}
		source.renameTo(dest);
	}

	private static void print(File transmittalFile, File employeeFile) throws PrintException, IOException {
		StringBuilder sb = new StringBuilder();
		
		BufferedReader reader = new BufferedReader(new FileReader(transmittalFile));
		String line = reader.readLine();
		while (line != null) {
			sb.append(line + "\n");
			line = reader.readLine();
		}
		reader.close();
		
		sb.append("\n");
		sb.append("\n");
		sb.append("\n");
		
		reader = new BufferedReader(new FileReader(employeeFile));
		line = reader.readLine();
		while (line != null) {
			sb.append(line + "\n");
			line = reader.readLine();
		}
		reader.close();
		sb.append("\f");
		
		PrintService service = PrintServiceLookup.lookupDefaultPrintService();

		// prints the famous hello world! plus a form feed
		InputStream is = new ByteArrayInputStream(sb.toString().getBytes("UTF8"));

		PrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
		pras.add(new Copies(2));

		DocFlavor flavor = DocFlavor.INPUT_STREAM.AUTOSENSE;
		Doc doc = new SimpleDoc(is, flavor, null);
		DocPrintJob job = service.createPrintJob();

		PrintJobWatcher pjw = new PrintJobWatcher(job);
		job.print(doc, pras);
		pjw.waitForDone();
		is.close();
	}

	static class PrintJobWatcher {
		boolean done = false;

		PrintJobWatcher(DocPrintJob job) {
			job.addPrintJobListener(new PrintJobAdapter() {
				public void printJobCanceled(PrintJobEvent pje) {
					allDone();
				}

				public void printJobCompleted(PrintJobEvent pje) {
					allDone();
				}

				public void printJobFailed(PrintJobEvent pje) {
					allDone();
				}

				public void printJobNoMoreEvents(PrintJobEvent pje) {
					allDone();
				}

				void allDone() {
					synchronized (PrintJobWatcher.this) {
						done = true;
						System.out.println("Printing done ...");
						PrintJobWatcher.this.notify();
					}
				}
			});
		}

		public synchronized void waitForDone() {
			try {
				while (!done) {
					wait();
				}
			} catch (InterruptedException e) {
			}
		}
	}

	private static Hashtable<String, Employee> getEmployees(XSSFSheet sheet, Employer employer,
			ApplicablePeriod applicablePeriod) {
		Hashtable<String, Employee> eeDB = Utility.getEEDB();
		for (int i = 9; i <= sheet.getLastRowNum(); i++) {
			XSSFRow row = sheet.getRow(i);

			if (row == null ||
				row.getCell(0) == null ||
				row.getCell(0).getRawValue().trim().equals("")) {
				break;
			}

			Employee employee = new Employee(row
				.getCell(0)
				.getStringCellValue()
				.replaceAll("-", ""), row.getCell(1).getStringCellValue(), row
				.getCell(2)
				.getStringCellValue(), employer, "0");
			employee.setMidInit(row.getCell(3).getStringCellValue());
			employee.setContribution(new Contribution(applicablePeriod, row
				.getCell(4)
				.getNumericCellValue(), row.getCell(5).getNumericCellValue(), row
				.getCell(6)
				.getCellType() == Cell.CELL_TYPE_STRING ? row.getCell(6).getStringCellValue() : row
				.getCell(6)
				.getRawValue()));

			eeDB.put(employee.getEeSssNumber(), employee);
		}

		Utility.updateEEDB(eeDB);
		return eeDB;
	}

}
