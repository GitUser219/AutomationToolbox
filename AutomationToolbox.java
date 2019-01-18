package automation.toolbox;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;
import javax.swing.AbstractAction;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import com.opencsv.CSVReader;

public class AutomationToolbox {
	private static XSSFSheet sheet;
	private static JButton generateFileButton;
	private static File xlsxFile;
	private static String imagesLocation;
	private static String password;
	private static String detailReportFilePath;
	private static String summaryReportFilePath;
	private static String fileDestination;

	public static void main(String[] args) {
		ImageIcon ready = new ImageIcon(ClassLoader.getSystemResource("images/ready.png"));
		ImageIcon notReady = new ImageIcon(ClassLoader.getSystemResource("images/not_ready.png"));
		ImageIcon redIcon = new ImageIcon(ClassLoader.getSystemResource("images/red_icon.png"));
		ImageIcon greenIcon = new ImageIcon(ClassLoader.getSystemResource("images/green_icon.png"));
		
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
			String message = "User interface not supported.";
			JOptionPane.showMessageDialog(null, message, "Failure", 0, notReady);
		}
		
		Font primaryFont = new Font("", Font.PLAIN, 16);
		JFrame mainFrame = new JFrame();
		mainFrame.setIconImage(new ImageIcon(ClassLoader.getSystemResource("images/black_icon.png")).getImage());
		mainFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		mainFrame.setTitle("Automation Toolbox");
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		mainFrame.setLocation(screenSize.width / 2 - 220, screenSize.height / 2 - 90);
		mainFrame.setSize(440, 180);
		mainFrame.setLayout(null);

		JLabel autoPdfIcon = new JLabel(new ImageIcon(redIcon.getImage().getScaledInstance(30, 30, java.awt.Image.SCALE_SMOOTH)));
		autoPdfIcon.setBounds(20, 20, 30, 30);
		mainFrame.add(autoPdfIcon);

		JLabel autoReportIcon = new JLabel(new ImageIcon(greenIcon.getImage().getScaledInstance(30, 30, java.awt.Image.SCALE_SMOOTH)));
		autoReportIcon.setBounds(20, 70, 30, 30);
		mainFrame.add(autoReportIcon);

		JButton pdfButton = new JButton(new AbstractAction("pdfAction") {
			private static final long serialVersionUID = 1L;
			
			@Override
			public void actionPerformed(ActionEvent e)
			{
				xlsxFile = new File("");
				imagesLocation = "";
				password = "";
				fileDestination = "";
				
				JFrame pdfFrame = new JFrame();
				pdfFrame.setIconImage(new ImageIcon(ClassLoader.getSystemResource("images/red_icon.png")).getImage());
				pdfFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
				pdfFrame.setTitle("PDF Creation");
				pdfFrame.setLocation(screenSize.width / 2 - 190, screenSize.height / 2 - 170);
				pdfFrame.setSize(380, 340);
				pdfFrame.setLayout(null);

				JLabel spreadsheetIcon = new JLabel(notReady);
				spreadsheetIcon.setBounds(305, 20, 30, 30);
				pdfFrame.add(spreadsheetIcon);
				
				JLabel imagesFolderIcon = new JLabel(notReady);
				imagesFolderIcon.setBounds(305, 70, 30, 30);
				pdfFrame.add(imagesFolderIcon);
				
				JLabel fileDestinationIcon = new JLabel(notReady);
				fileDestinationIcon.setBounds(305, 120, 30, 30);
				pdfFrame.add(fileDestinationIcon);

				JLabel passwordLabel = new JLabel("Password protect the file");
				passwordLabel.setFont(primaryFont);
				passwordLabel.setHorizontalAlignment(JLabel.CENTER);
				passwordLabel.setBounds(17, 160, 315, 30);
				pdfFrame.add(passwordLabel);
				
				JTextField passwordField = new JTextField();
				passwordField.setFont(primaryFont);
				passwordField.setBounds(20, 190, 315, 30);
				pdfFrame.add(passwordField);
				
				ArrayList<String[]> xlsxData = new ArrayList<String[]>();

				JButton xlsxButton = new JButton(new AbstractAction("xlsxAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent event) {
						JFileChooser xlsxFileChooser = new JFileChooser();
						xlsxFileChooser.setDialogTitle("Browse for the spreadsheet");
						xlsxFileChooser.setFileFilter(new FileNameExtensionFilter("Microsoft Excel XLSX", "xlsx"));

						if (xlsxFileChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
							if (xlsxFileChooser.getSelectedFile().exists()) {
								xlsxFile = xlsxFileChooser.getSelectedFile();
								spreadsheetIcon.setIcon(ready);
								
								if (!imagesLocation.equals("") && !fileDestination.equals("")) {
									generateFileButton.setEnabled(true);
								}
							} else {
								JOptionPane.showMessageDialog(null, "Invalid selection.", "Failure", 0, notReady);
							}
						}
					}
				});
				xlsxButton.setText("Select spreadsheet");
				xlsxButton.setFont(primaryFont);
				xlsxButton.setBounds(20, 20, 265, 30);
				pdfFrame.add(xlsxButton);
				
				JButton imagesButton = new JButton(new AbstractAction("imagesAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent event) {
						JFileChooser imagesFolderChooser = new JFileChooser();
						imagesFolderChooser.setDialogTitle("Browse for the images folder");
						imagesFolderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

						if (imagesFolderChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
							if (imagesFolderChooser.getSelectedFile().exists()) {
								imagesLocation = imagesFolderChooser.getSelectedFile().getPath() + "\\";
								imagesFolderIcon.setIcon(ready);
	
								if (xlsxFile.exists() && !fileDestination.equals("")) {
									generateFileButton.setEnabled(true);
								}
							} else {
								JOptionPane.showMessageDialog(null, "Invalid selection.", "Failure", 0, notReady);
							}
						}
					}
				});
				imagesButton.setText("Select images location");
				imagesButton.setFont(primaryFont);
				imagesButton.setBounds(20, 70, 265, 30);
				pdfFrame.add(imagesButton);

				JButton destinationButton = new JButton(new AbstractAction("destinationAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent event) {
						JFileChooser destinationChooser = new JFileChooser();
						destinationChooser.setDialogTitle("Select the destination");
						destinationChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

						if (destinationChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
							if (destinationChooser.getSelectedFile().exists()) {
								fileDestinationIcon.setIcon(ready);
								fileDestination = destinationChooser.getSelectedFile().getPath() + "\\";
	
								if (xlsxFile.exists() && !imagesLocation.equals("")) {
									generateFileButton.setEnabled(true);
								}
							} else {
								JOptionPane.showMessageDialog(null, "Invalid selection.", "Failure", 0, notReady);
							}
						}
					}
				});
				destinationButton.setText("Select destination");
				destinationButton.setFont(primaryFont);
				destinationButton.setBounds(20, 120, 265, 30);
				pdfFrame.add(destinationButton);

				generateFileButton = new JButton(new AbstractAction("generateFileAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent e) {
						password = passwordField.getText();

						if (password.equals("")) {
							String title = "Password required";
							String body = "A password is required to protect the file.";
							JOptionPane.showMessageDialog(null, body, title, 0, notReady);
						} else {
							try {
								FileInputStream inputStream = new FileInputStream(xlsxFile);
								XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
								sheet = workbook.getSheetAt(0);
								workbook.close();
								
								if (sheet.getLastRowNum() < 2) {
									throw new Exception();
								}
								
								for (int rowIndex = 1; rowIndex < sheet.getLastRowNum()+1; rowIndex++) {
									String[] row = new String[26];

									for (int cellIndex = 0; cellIndex < 26; cellIndex++) {
										String cellValue = sheet.getRow(rowIndex).getCell(cellIndex).toString();
										
										if (cellIndex > 12 && cellIndex < 19) {
											cellValue = cellValue.toUpperCase();
											
											switch (cellValue) {
												case "YES":
												case "NO":
												case "N/A":
													row[cellIndex] = cellValue;
													break;
												default:
													throw new Exception();
											}
										} else {
											row[cellIndex] = cellValue;
										}
									}

									xlsxData.add(row);
								}
							} catch (Exception ex) {
								String body = "Failed to read the data from the spreadsheet.\n\n"
										+ "Make sure that the following statements are true:\n"
										+ "1. There are at least two rows in the spreadsheet"
										+ "2. Each row contains exactly 26 non-empty cells"
										+ "3. Columns 13 to 18 contain \"Yes\", \"No\", or \"N/A\""
										+ "4. The spreadsheet is not open in another program\"\n\n"
										+ "If you continue to have issues, contact Jonathan Nowak at\n"
										+ "jnowak@quantumcrossings.com or (217) 776-9115";;
								JOptionPane.showMessageDialog(null, body, "Failure", 0, notReady);
							}

							String fileName = xlsxData.get(1)[22] + ".pdf";

							try {
								ArrayList<String> logs = new ArrayList<String>();
								Document pdf = new Document(PageSize.A4.rotate());
								PdfWriter writer = PdfWriter.getInstance(pdf, new FileOutputStream(fileDestination + fileName));
								writer.setCompressionLevel(9);
								writer.setEncryption(password.getBytes(), password.getBytes(), PdfWriter.ALLOW_PRINTING, PdfWriter.STANDARD_ENCRYPTION_128);
								pdf.open();

								for (int j = 0; j < xlsxData.size(); j++) {
									String[] row = xlsxData.get(j);
									pdf.newPage();
									PdfContentByte cb = writer.getDirectContentUnder();
									Image background = Image.getInstance(ClassLoader.getSystemResource("images/background.jpg"));
									background.scaleAbsolute(840.0f, 595.0f);
									background.setAbsolutePosition(0, 0);
									cb.addImage(background);
									cb = writer.getDirectContent();
									BaseFont baseFont = BaseFont.createFont();
									cb.setFontAndSize(baseFont, 12);
									float X = 93.0f, Y = 481.0f, H = 101.0f;
									String[] titles = {"1", "2", "3", "4", "5", "6", "Before", "After", "ROI"};

									for (int i = 0; i < 35; i++) {
										switch (i) {
											case 1:  X = 125.0f; Y = 463.0f; break;
											case 2:  X = 131.0f; Y = 443.0f; break;
											case 3:  X = 96.0f;  Y = 423.0f; break;
											case 4:  X = 151.0f; Y = 403.0f; break;
											case 5:  X = 161.0f; Y = 383.0f; break;
											case 6:  X = 198.0f; Y = 363.0f; break;
											case 7:  X = 295.0f; Y = 480.0f; break;
											case 8:  X = 309.0f; Y = 458.0f; break;
											case 9:  X = 319.0f; Y = 438.0f; break;
											case 10: X = 317.0f; Y = 419.0f; break;
											case 11: X = 354.0f; Y = 400.0f; break;
											case 12: X = 282.0f; Y = 381.0f; break;
											case 13:             Y = 298.0f; break;
											case 14:             Y = 281.0f; break;
											case 15:             Y = 265.0f; break;
											case 16:             Y = 247.0f; break;
											case 17:             Y = 231.0f; break;
											case 18:             Y = 214.0f; break;
											case 19: X = 60.0f;  Y = 145.0f; break;
											case 20: X = 151.0f; Y = 45.0f;  break;
											case 21: X = 345.0f;             break;
											case 22: X = 525.0f; Y = 143.0f; break;
											case 23: X = 535.0f; Y = 122.0f; break;
											case 24: X = 502.0f; Y = 103.0f; break;
											case 25: X = 580.0f; Y = 83.0f;  break;
											case 26: X = 453.0f; Y = 395.5f; break;
											case 27:             Y = 285.5f; break;
											case 28:             Y = 176.5f; break;
											case 29: X = 583.0f;             break;
											case 30:             Y = 285.5f; break;
											case 31:             Y = 395.5f; break;
											case 32: X = 713.0f; H = 84.0f;  break;
											case 33:             Y = 285.5f; break;
											case 34:             Y = 176.5f; break;
										}

										if (i > 25) {
											String imageFileName = (row[22] + " " + row[0] + " " + titles[i - 26]);
											String extension = "";

											if ((new File(imagesLocation + imageFileName + ".jpg")).exists()) {
												extension = ".jpg";
											} else if ((new File(imagesLocation + imageFileName + ".jpeg")).exists()) {
												extension = ".jpeg";
											} else if ((new File(imagesLocation + imageFileName + ".png")).exists()) {
												extension = ".png";
											}

											if (extension.equals("")) {
												logs.add("Image not found: " + imageFileName + "\n");
											} else {
												Image image = Image.getInstance(imagesLocation + imageFileName + extension);
												image.scaleAbsolute(119.0f, H);
												image.setAbsolutePosition(X, Y);
												cb.addImage(image);
											}
										} else {
											cb.beginText();
						
											if (i > 12 && i < 19) {
												switch (row[i]) {
													case "YES":
														X = 268.0f;
														break;
													case "NO":
														X = 303.0f;
														break;
													case "N/A":
														X = 336.0f;
														break;
												}
												
												cb.moveText(X, Y);
												cb.showText("X");
											} else {
												cb.moveText(X, Y);
												cb.showText(row[i]);
											}
											
											cb.endText();
										}
									}
								}
								
								pdf.close();

								if (logs.size() > 0) {
									PrintWriter logFile = new PrintWriter(fileDestination + xlsxData.get(1)[22] + " PDF Creation Log.txt", "UTF-8");

									for (String log : logs) {
										logFile.println(log);
									}

									logFile.close();

									pdfFrame.setVisible(false);
									pdfFrame.dispose();

									String body = xlsxData.get(1)[22]
											+ ".pdf was successfully created, but there are images missing.\n\n"
											+ "Check " + xlsxData.get(1)[22]
											+ " PDF Creation Log.txt for more information.";
									JOptionPane.showMessageDialog(null, body, "Success", 0, ready);
								} else {
									pdfFrame.setVisible(false);
									pdfFrame.dispose();
									
									String body = xlsxData.get(1)[22] + ".pdf was successfully created.";
									JOptionPane.showMessageDialog(null, body, "Success", 0, ready);
								}
							} catch (Exception ex) {
								String title = "Failed to write to file: " + fileName;
								String body = "Make sure that " + fileName + " is not open in another program.\"\n\n"
										+ "If you continue to have issues, contact Jonathan Nowak at\n"
										+ "jnowak@quantumcrossings.com or (217) 776-9115";
								JOptionPane.showMessageDialog(null, body, title, 0, notReady);
							}
						}
					}
				});
				generateFileButton.setText("Generate the file");
				generateFileButton.setFont(primaryFont);
				generateFileButton.setBounds(20, 240, 315, 30);
				generateFileButton.setEnabled(false);
				pdfFrame.add(generateFileButton);
				pdfFrame.setVisible(true);
			}
		});
		pdfButton.setText("Preventative Maintenance PDF Creation");
		pdfButton.setBounds(70, 20, 330, 30);
		pdfButton.setFont(primaryFont);
		pdfButton.setSelected(false);
		mainFrame.add(pdfButton);

		JButton reportButton = new JButton(new AbstractAction("reportAction") {
			private static final long serialVersionUID = 1L;

			@Override
			public void actionPerformed(ActionEvent e) {
				detailReportFilePath = "";
				summaryReportFilePath = "";
				fileDestination = "";
				
				JFrame reportFrame = new JFrame();
				reportFrame.setIconImage(new ImageIcon(ClassLoader.getSystemResource("images/green_icon.png")).getImage());
				reportFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
				reportFrame.setTitle("Report Creation");
				reportFrame.setLocation(screenSize.width / 2 - 200, screenSize.height / 2 - 135);
				reportFrame.setSize(400, 270);
				reportFrame.setLayout(null);

				JLabel detailReportIcon = new JLabel(notReady);
				detailReportIcon.setBounds(325, 20, 30, 30);
				reportFrame.add(detailReportIcon);

				JLabel summaryReportIcon = new JLabel(notReady);
				summaryReportIcon.setBounds(325, 70, 30, 30);
				reportFrame.add(summaryReportIcon);

				JLabel fileDestinationIcon = new JLabel(notReady);
				fileDestinationIcon.setBounds(325, 120, 30, 30);
				reportFrame.add(fileDestinationIcon);

				JButton detailReportSelectButton = new JButton(new AbstractAction("detailReportSelectAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent event) {
						JFileChooser detailReportChooser = new JFileChooser();
						detailReportChooser.setDialogTitle("Browse for the detail report");
						detailReportChooser.setFileFilter(new FileNameExtensionFilter("CSV File", "csv"));

						if (detailReportChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
							if (detailReportChooser.getSelectedFile().exists()) {
								detailReportFilePath = detailReportChooser.getSelectedFile().getPath();
								detailReportIcon.setIcon(ready);
	
								if (!summaryReportFilePath.equals("") && !fileDestination.equals("")) {
									generateFileButton.setEnabled(true);
								}
							} else {
								JOptionPane.showMessageDialog(null, "Invalid selection.", "Failure", 0, notReady);
							}
						}
					}
				});
				detailReportSelectButton.setText("Select the detail report");
				detailReportSelectButton.setBounds(20, 20, 285, 30);
				detailReportSelectButton.setFont(primaryFont);
				reportFrame.add(detailReportSelectButton);

				JButton summaryReportSelectButton = new JButton(new AbstractAction("summaryReportSelectAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent event) {
						JFileChooser summaryReportChooser = new JFileChooser();
						summaryReportChooser.setDialogTitle("Browse for the summary report");
						summaryReportChooser.setFileFilter(new FileNameExtensionFilter("CSV File", "csv"));

						if (summaryReportChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
							if (summaryReportChooser.getSelectedFile().exists()) {
								summaryReportFilePath = summaryReportChooser.getSelectedFile().getPath();
								summaryReportIcon.setIcon(ready);
	
								if (!detailReportFilePath.equals("") && !fileDestination.equals("")) {
									generateFileButton.setEnabled(true);
								}
							} else {
								JOptionPane.showMessageDialog(null, "Invalid selection.", "Failure", 0, notReady);
							}
						}
					}
				});
				summaryReportSelectButton.setText("Select the summary report");
				summaryReportSelectButton.setBounds(20, 70, 285, 30);
				summaryReportSelectButton.setFont(primaryFont);
				reportFrame.add(summaryReportSelectButton);

				JButton fileDestinationSelectButton = new JButton(new AbstractAction("fileDestinationSelectAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent event) {
						JFileChooser fileDestinationChooser = new JFileChooser();
						fileDestinationChooser.setDialogTitle("Browse for the file destination");
						fileDestinationChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

						if (fileDestinationChooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
							if (fileDestinationChooser.getSelectedFile().exists()) {
								fileDestination = fileDestinationChooser.getSelectedFile().getPath() + "\\Report.xlsx";
								fileDestinationIcon.setIcon(ready);
								
								if (!detailReportFilePath.equals("") && !summaryReportFilePath.equals("")) {
									generateFileButton.setEnabled(true);
								}
							} else {
								JOptionPane.showMessageDialog(null, "Invalid selection.", "Failure", 0, notReady);
							}
						}
					}
				});
				fileDestinationSelectButton.setText("Select the file destination");
				fileDestinationSelectButton.setBounds(20, 120, 285, 30);
				fileDestinationSelectButton.setFont(primaryFont);
				reportFrame.add(fileDestinationSelectButton);

				generateFileButton = new JButton(new AbstractAction("generateFileAction") {
					private static final long serialVersionUID = 1L;

					@Override
					public void actionPerformed(ActionEvent e) {
						try {
							ArrayList<String[]> reportData = new ArrayList<String[]>();
							CSVReader csvReader = new CSVReader(Files.newBufferedReader(Paths.get(summaryReportFilePath)));

							while (csvReader.peek() != null) {
								String[] summaryReportRow = csvReader.readNext();
								String[] reportDataRow = new String[11];

								if (!summaryReportRow[12].equals("QC")) {
									reportDataRow[0] = summaryReportRow[0];
									reportDataRow[1] = summaryReportRow[2];
									reportDataRow[2] = summaryReportRow[6];
									reportDataRow[3] = summaryReportRow[10];
									reportDataRow[4] = summaryReportRow[12];
									reportDataRow[5] = summaryReportRow[13];
									reportDataRow[6] = summaryReportRow[15];
									reportDataRow[7] = summaryReportRow[17];
									reportDataRow[8] = summaryReportRow[39];

									reportData.add(reportDataRow);
								}
							}

							csvReader.close();

							if (!reportData.get(0)[8].equals("Initial Work Complete Comment")) {
								throw new Exception();
							}

							ArrayList<String[]> detailData = new ArrayList<String[]>();
							csvReader = new CSVReader(Files.newBufferedReader(Paths.get(detailReportFilePath)));

							while (csvReader.peek() != null) {
								String[] detailReportRow = csvReader.readNext();
								String[] detailDataRow = new String[4];

								detailDataRow[0] = detailReportRow[1];
								detailDataRow[1] = detailReportRow[4];
								detailDataRow[2] = detailReportRow[13];
								detailDataRow[3] = detailReportRow[15];

								detailData.add(detailDataRow);
							}

							csvReader.close();

							if (!detailData.get(0)[3].equals("Comments")) {
								throw new Exception();
							}

							for (int i = 0; i < reportData.size(); i++) {
								reportData.get(i)[9] = "N/A";
								reportData.get(i)[10] = "N/A";
								reportData.get(i)[7] = "[Initial]\n" + reportData.get(i)[7];

								for (int j = 0; j < detailData.size(); j++) {
									if (reportData.get(i)[1].equals(detailData.get(j)[1])) {
										reportData.get(i)[9] = detailData.get(j)[0];
										String update = detailData.get(j)[3];
										
										if (!detailData.get(j)[2].equals("Initial") && !update.equals("Work Order printed.")) {
											reportData.get(i)[7] = reportData.get(i)[7] + "\n\n[" + detailData.get(j)[2] + "]\n" + update;

											if (update.contains("SCHD ") && update.substring(update.indexOf("SCHD ") + 5).length() > 9) {
												reportData.get(i)[10] = update.substring(update.indexOf("SCHD ") + 5, update.indexOf("SCHD ") + 15);
											}
										}

										detailData.remove(j);
										j--;
									}
								}
							}

							XSSFWorkbook workbook = new XSSFWorkbook();
							sheet = workbook.createSheet("Report");
							sheet.setDefaultRowHeight(new Short("400"));
							sheet.createFreezePane(0, 1);

							CellStyle headerStyle = workbook.createCellStyle();
							XSSFFont headerFont = workbook.createFont();
							headerFont.setBold(true);
							headerStyle.setFont(headerFont);
							headerStyle.setAlignment(HorizontalAlignment.CENTER);
							headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
							headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
							headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							headerStyle.setBorderRight(BorderStyle.THIN);
							headerStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

							String month = Integer.toString(LocalDate.now().getMonthValue());
							String day = Integer.toString(LocalDate.now().getDayOfMonth());
							String year = Integer.toString(LocalDate.now().getYear());
							String today = month + "/" + day + "/" + year;
							
							Row headerRow = sheet.createRow(0);

							for (int i = 0; i < 16; i++) {
								Cell headerCell = headerRow.createCell(i);
								String headerCellValue = "";

								switch (i) {
									case 0: headerCellValue = "OpCo";                break;
									case 1: headerCellValue = "Work Status";         break;
									case 2: headerCellValue = "JIRA Status";         break;
									case 3: headerCellValue = "Site Type";           break;
									case 4: headerCellValue = "Location";            break;
									case 5: headerCellValue = "System";              break;
									case 6: headerCellValue = "SIT Number";          break;
									case 7: headerCellValue = "Date Received";       break;
									case 8: headerCellValue = "Issue";               break;
									case 9: headerCellValue = "Date Completed";      break;
									case 10: headerCellValue = "Days Open";          break;
									case 11: headerCellValue = "Comments";           break;
									case 12: headerCellValue = "Request ID";         break;
									case 13: headerCellValue = "Urgency";            break;
									case 14: headerCellValue = "Scheduled Date";     break;
									case 15: headerCellValue = "Run Date: " + today; break;
								}

								headerCell.setCellValue(headerCellValue);
								headerCell.setCellStyle(headerStyle);
							}

							if (month.length() == 1) {
								month = "0" + month;
							}
							if (day.length() == 1) {
								day = "0" + day;
							}
							today = month + "/" + day + "/" + year;

							CellStyle primaryStyle = workbook.createCellStyle();
							primaryStyle.setAlignment(HorizontalAlignment.CENTER);
							primaryStyle.setVerticalAlignment(VerticalAlignment.CENTER);
							
							for (int i = 1; i < reportData.size(); i++) {
								Row reportRow = sheet.createRow(i);

								String opCo;
								String workStatus = reportData.get(i)[4];
								String jiraStatus;
								String siteType;
								String location = reportData.get(i)[2];
								String system = reportData.get(i)[3];
								String sitNumber;
								String dateOpened = reportData.get(i)[0];
								String issue = reportData.get(i)[7];
								String dateCompleted;
								Long daysOpen;
								String comments = reportData.get(i)[8];
								int requestID = Integer.parseInt(reportData.get(i)[1]);
								String urgency;
								String scheduledDate = reportData.get(i)[10];

								if (reportData.get(i)[9].contains("-")) {
									opCo = reportData.get(i)[9].substring(0, reportData.get(i)[9].indexOf("-") - 1);
									siteType = reportData.get(i)[9].substring(reportData.get(i)[9].indexOf("-") + 2);
								} else {
									opCo = "N/A";
									siteType = "N/A";
								}

								if (reportData.get(i)[4].equals("Closed")) {
									jiraStatus = "Closed";
								} else {
									jiraStatus = "Open";
								}

								if (issue.indexOf("SIT-") != -1) {
									int index = issue.indexOf("SIT-");
									char[] sit = issue.substring(index).toCharArray();

									for (int c = 4; c < sit.length; c++) {
										if (sit[c] != '0' && sit[c] != '1' && sit[c] != '2' && sit[c] != '3'
														  && sit[c] != '4' && sit[c] != '5' && sit[c] != '6'
														  && sit[c] != '7' && sit[c] != '8' && sit[c] != '9') {
											index += c;
											c = sit.length;
										}
									}

									sitNumber = issue.substring(issue.indexOf("SIT-"), index);
								} else {
									sitNumber = "N/A";
								}

								if (reportData.get(i)[5].equals("Low")) {
									urgency = "Routine";
								} else if (reportData.get(i)[5].equals("Medium")) {
									urgency = "Urgent";
								} else if (reportData.get(i)[5].equals("High")) {
									urgency = "Emergency";
								} else {
									urgency = "N/A";
								}

								if (reportData.get(i)[6].indexOf(" ") != -1) {
									dateCompleted = reportData.get(i)[6].substring(0, reportData.get(i)[6].indexOf(" "));
								} else {
									dateCompleted = "Not complete";
								}

								SimpleDateFormat f = new SimpleDateFormat("MM/dd/yyyy");
								
								if (dateCompleted.equals("Not complete")) {
									daysOpen = TimeUnit.DAYS.convert(f.parse(today).getTime() - f.parse(dateOpened).getTime(), TimeUnit.MILLISECONDS);
								} else {
									daysOpen = TimeUnit.DAYS.convert(f.parse(dateCompleted).getTime() - f.parse(dateOpened).getTime(), TimeUnit.MILLISECONDS);
								}

								for (int j = 0; j < 15; j++) {
									Cell cell = reportRow.createCell(j);
									cell.setCellStyle(primaryStyle);

									switch (j) {
										case 0: cell.setCellValue(opCo);           break;
										case 1: cell.setCellValue(workStatus);     break;
										case 2: cell.setCellValue(jiraStatus);     break;
										case 3: cell.setCellValue(siteType);       break;
										case 4: cell.setCellValue(location);       break;
										case 5: cell.setCellValue(system);         break;
										case 6: cell.setCellValue(sitNumber);      break;
										case 7: cell.setCellValue(dateOpened);     break;
										case 8: cell.setCellValue(issue);          break;
										case 9: cell.setCellValue(dateCompleted);  break;
										case 10: cell.setCellType(CellType.NUMERIC);
											     cell.setCellValue(daysOpen);      break;
										case 11: cell.setCellValue(comments);      break;
										case 12: cell.setCellType(CellType.NUMERIC);
											     cell.setCellValue(requestID);     break;
										case 13: cell.setCellValue(urgency);       break;
										case 14: cell.setCellValue(scheduledDate); break;
									}
								}
							}

							for (int i = 0; i < 16; i++) {
								if (i == 8) {
									sheet.setColumnWidth(8, 1800);
								} else if (i == 11) {
									sheet.setColumnWidth(11, 3000);
								} else {
									sheet.autoSizeColumn(i);
								}
							}

							FileOutputStream outputStream = new FileOutputStream(fileDestination);
							workbook.write(outputStream);
							outputStream.close();
							workbook.close();

							reportFrame.setVisible(false);
							reportFrame.dispose();

							String message = "Report.xlsx was successfully created.";
							JOptionPane.showMessageDialog(null, message, "Success", 0, ready);
						} catch (Exception ex) {
							String message = "Failed to generate the file.\n\n"
									+ "Make sure that the following statements are true:\n"
									+ "1. Report.xlsx is not open in another program\n"
									+ "2. The detail report is not open in another program\n"
									+ "3. The summary report is not open in another program\n"
									+ "4. The CSV files are raw exports from 360 Facility\n"
									+ "5. The CSV files have not been modified by another program\n"
									+ "6. You have selected the correct CSV files\n"
									+ "7. The scheduled date format is \"SCHD MM/DD/YYYY\"\n\n"
									+ "If you continue to have issues, contact Jonathan Nowak at\n"
									+ "jnowak@quantumcrossings.com or (217) 776-9115";
							JOptionPane.showMessageDialog(null, message, "Failure", 0, notReady);
						}
					}
				});
				generateFileButton.setText("Generate the file");
				generateFileButton.setBounds(20, 170, 335, 30);
				generateFileButton.setFont(primaryFont);
				generateFileButton.setEnabled(false);
				reportFrame.add(generateFileButton);
				reportFrame.setVisible(true);
			}
		});
		reportButton.setText("Preventative Maintenance Report Creation");
		reportButton.setBounds(70, 70, 330, 30);
		reportButton.setFont(primaryFont);
		mainFrame.add(reportButton);
		mainFrame.setVisible(true);
	}
}