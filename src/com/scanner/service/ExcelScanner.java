package com.scanner.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;
import java.util.TimeZone;

import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.FillFormat;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.License;
import com.aspose.cells.MsoPresetTextEffect;
import com.aspose.cells.PlacementType;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExcelScanner {

	double sumRowHeightCoveredByWatermark = 0;
	double sumColumnsWidthCoveredByWatermark = 0;
	private int sheetWaterMarkCount;

	public static void main(String[] args) throws FileNotFoundException {

		ExcelScanner excelScanner = new ExcelScanner();
		excelScanner.applyLicense();

		Scanner sc = new Scanner(System.in);

		// Source Path
		String sourcePath = readString(sc, "Enter your Source Path");

		// Target Path
		String targetPath = readString(sc, "Enter your Target Path");

		// Water mark Text
		String watermarkText = readString(sc, "Enter your Watermark Text");

		System.out.println("Please select your sheet Index");

		File file = new File(sourcePath);
		byte[] xFile = new byte[(int) file.length()];
		FileInputStream fis;
		try {
			fis = new FileInputStream(file);
			fis.read(xFile); // read file into bytes[]
			fis.close();
			final Date currentTime = new Date();

			final SimpleDateFormat sdf = new SimpleDateFormat("EEE, MMM d, yyyy hh:mm:ss a z");

			// Give it to me in GMT time.
			sdf.setTimeZone(TimeZone.getTimeZone("GMT"));

			String fullWatermarkText = watermarkText + System.lineSeparator() + "userMerill@merrillcorp.com"
					+ System.lineSeparator() + sdf.format(currentTime) + "(Coordinated Universal Time)";

			excelScanner.displaySheets(xFile);

			// Sheet index
			int totalCount = readInt(sc, "Enter total number of sheet to apply watermark");
			
			int[] sheetIndex = new int[totalCount];

			for (int index = 0; index < totalCount; index++) {
				sheetIndex[index] = readInt(sc, "Enter the sheet index");
			}
			excelScanner.addWatermarkToExcelWorkbook(xFile, fullWatermarkText, targetPath, sheetIndex, sc);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			sc.close();
		}

	}

	private static String readString(Scanner sc, String msg) {
		String value = null;
		do {
			System.out.println(msg);
			value = sc.nextLine();
			if (value.isEmpty()) {
				System.out.println("Nothing was entered. Please try again");
			}
		} while (value.isEmpty());
		return value;
	}
	
	private static int readInt(Scanner sc, String msg) {
		int number;
		do {
			System.out.println(msg);
			if (!sc.hasNextInt()) {
				System.out.println("That's not a number!");
				sc.next(); 
			}
			number = sc.nextInt();
		} while (number <= 0);
		return number;
	}

	private void applyLicense() {
		try (InputStream inputStream = new FileInputStream(
				new File(this.getClass().getResource("/Aspose.Cells.lic").getFile()))) {
			License license = new License();
			license.setLicense(inputStream);

		} catch (Exception e) {

		}
	}

	private void displaySheets(byte[] xFile) {
		Workbook workbook;
		try {
			workbook = new Workbook(new ByteArrayInputStream(xFile));
			for (int sheetIndex = 0; sheetIndex < workbook.getWorksheets().getCount(); sheetIndex++) {
				Worksheet sheet = workbook.getWorksheets().get(sheetIndex);
				System.out.println("SheetIndex:::" + (sheetIndex + 1) + " with SheetName:::" + sheet.getName());
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void addWatermarkToExcelWorkbook(byte[] xFile, String watermarkText, String targetPath, int[] sheetIndex,
			Scanner sc) {

		Workbook workbook;
		Worksheet sheet;
		int watermarkCount = 0;
		char sheetData;

		try {
			workbook = new Workbook(new ByteArrayInputStream(xFile));

			for (int index = 0; index < sheetIndex.length; index++) {
				sheetWaterMarkCount = 0;
				sheet = workbook.getWorksheets().get(sheetIndex[index] - 1);
				// Applying water mark to sheets
				System.out.println("Do you want the watermark for the Entire sheet or Available data(E/A) for sheet Number:::"
						+ sheetIndex[index]);

				do {
					sheetData = sc.next().charAt(0);
					if (!(sheetData == 'E' || sheetData == 'A')) {
						System.out
								.println("Wrong Input, Pls enter E for Entire Sheet or A for Available data for sheet Number:::"
										+ sheetIndex[index]);
					}
				} while (!(sheetData == 'E' || sheetData == 'A'));
				if (sheetData == 'E') {
					do {
						// Water mark Count
						watermarkCount = readInt(sc, "Enter your Watermark Count");						 

						if (watermarkCount > 5000) {
							System.out.println("Please enter not more than 5000");
						}
					} while (watermarkCount > 5000);

					if (watermarkCount == 0) {
						System.out.println(
								"Alert::::Watermark can be applied with Available data, Please select 'A' option below");
					}
					if (isEmptySheet(sheet)) {
						addWordArtForTileWordArtForEmptySheet(sheet, watermarkText, watermarkCount);
					} else {
						addWordArtForTileWordArt(sheet, watermarkText, watermarkCount);
					}
				} else if (sheetData == 'A') {
					addWordArtForTileWordArtV2(sheet, watermarkText, watermarkCount);
				}
			}

			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			workbook.save(outputStream, FileFormatType.XLSX);
			byte[] buffer = new byte[(int) outputStream.size()];
			buffer = outputStream.toByteArray();
			FileOutputStream fileOut = new FileOutputStream(targetPath);
			System.out.println("Done changes::::");
			fileOut.write(buffer);
			fileOut.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private boolean isEmptySheet(Worksheet sheet) {
		Cells cells = sheet.getCells();
		if (cells.getMaxDataRow() == -1 && cells.getMaxDataColumn() == -1)
			return true;
		else
			return false;
	}

	private void addWordArtForTileWordArtForEmptySheet(Worksheet sheet, String watermarkText, int watermarkCount) {
		int printRow = 0;
		int printColumn = 0;
		int maxColumnCount = 250;

		if (watermarkCount == 0) {
			System.out.println("Alert::: Watermark count is zero, so not applying any watermark");
			return;
		}

		for (int row = 0; row < watermarkCount; row++) {
			Shape wordArt = sheet.getShapes().addTextEffect(MsoPresetTextEffect.TEXT_EFFECT_1, watermarkText, "Arial",
					50, false, true, printRow, 8, printColumn, 1, 130, 600);
			wordArt.setRotationAngle(-40);
			wordArt.setPlacement(PlacementType.FREE_FLOATING);
			FillFormat wordArtFormat = wordArt.getFill();
			wordArtFormat.setFillType(3);
			wordArtFormat.setOneColorGradient(Color.getGray(), 0.0, GradientStyleType.HORIZONTAL, 2);
			wordArtFormat.setTransparency(0.8);
			sheetWaterMarkCount++;

			if (watermarkCount <= maxColumnCount && row == watermarkCount / 2) {
				printRow += 20;
				printColumn = 0;
			} else if (row > maxColumnCount && row % maxColumnCount == 0) {
				printRow += 20;
				printColumn = 0;
			} else {
				printColumn += 10;
			}

		}
		System.out.println("sheetWaterMarkCount==>" + sheetWaterMarkCount);

	}

	private void addWordArtForTileWordArt(Worksheet sheet, String watermarkText, int watermarkCount) {
		System.out.println("addWordArtForTileWordArt() with sheetName::::" + sheet.getName());
		Cells cells = sheet.getCells();
		int watermarkWidth = 600;
		double sumRowHeightCoveredByWatermark = 0;
		double sumColumnsWidthCoveredByWatermark = 0;
		int printRow = 0;

		for (int row = 0; row < watermarkCount; row++) {
			sumRowHeightCoveredByWatermark += cells.getRowHeightPixel(row);
			if ((sumRowHeightCoveredByWatermark > watermarkWidth || row == 1) && sheetWaterMarkCount < watermarkCount) {
				for (int column = 1; column < watermarkCount; column++) {
					sumColumnsWidthCoveredByWatermark += cells.getColumnWidthPixel(column);
					if ((sumColumnsWidthCoveredByWatermark > watermarkWidth || column == 1)
							&& sheetWaterMarkCount < watermarkCount) {
						Shape wordArt = sheet.getShapes().addTextEffect(MsoPresetTextEffect.TEXT_EFFECT_1,
								watermarkText, "Arial", 50, false, true, printRow, 8, column, 1, 130, watermarkWidth);
						wordArt.setRotationAngle(-40);
						wordArt.setPlacement(PlacementType.FREE_FLOATING);
						FillFormat wordArtFormat = wordArt.getFill();
						wordArtFormat.setFillType(3);
						wordArtFormat.setOneColorGradient(Color.getGray(), 0.0, GradientStyleType.HORIZONTAL, 2);
						wordArtFormat.setTransparency(0.8);
						sumColumnsWidthCoveredByWatermark = 0;
						sheetWaterMarkCount++;
					}
				}
				printRow += 20;
			}
		}
		System.out.println("sheetWaterMarkCount==>" + sheetWaterMarkCount);

	}

	private void addWordArtForTileWordArtV2(Worksheet sheet, String watermarkText, int watermarkCount) {
		System.out.println("addWordArtForTileWordArtV2() with sheetName::::" + sheet.getName());
		Cells cells = sheet.getCells();
		int watermarkWidth = 500;
		int watermarkHeight = 90;
		double sumRowHeightCoveredByWatermark = 0;
		double sumColumnsWidthCoveredByWatermark = 0;
		int degree = -40;
		double watermarkBoundingBoxHeight = Math.abs(watermarkWidth * Math.sin(Math.toRadians(degree)))
				+ Math.abs(watermarkHeight * Math.cos(Math.toRadians(degree)));
		double watermarkBoundingBoxWidth = Math.abs(watermarkWidth * Math.cos(Math.toRadians(degree)))
				+ Math.abs(watermarkHeight * Math.sin(Math.toRadians(degree)));
		sumColumnsWidthCoveredByWatermark = watermarkBoundingBoxWidth;
		sumRowHeightCoveredByWatermark = watermarkBoundingBoxHeight;

		for (int upperLeftRow = 0; upperLeftRow <= cells.getMaxDataRow(); upperLeftRow++) {
			if (sumRowHeightCoveredByWatermark >= watermarkBoundingBoxHeight) {
				for (int upperLeftColumn = 0; upperLeftColumn <= cells.getMaxDataColumn(); upperLeftColumn++) {
					if (sumColumnsWidthCoveredByWatermark >= watermarkBoundingBoxWidth) {

						Shape wordArt = sheet.getShapes().addTextEffect(MsoPresetTextEffect.TEXT_EFFECT_1,
								watermarkText, "Arial Black", 50, false, true, upperLeftRow, 8, upperLeftColumn, 1,
								watermarkHeight, watermarkWidth);
						wordArt.setRotationAngle(degree);
						wordArt.setPlacement(PlacementType.FREE_FLOATING);
						FillFormat wordArtFormat = wordArt.getFill();
						wordArtFormat.setFillType(3);
						wordArtFormat.setOneColorGradient(Color.getGray(), 0.0, GradientStyleType.HORIZONTAL, 2);
						wordArtFormat.setTransparency(0.8);

						sumColumnsWidthCoveredByWatermark = 0;
						sheetWaterMarkCount++;
					}
					sumColumnsWidthCoveredByWatermark += cells.getColumnWidthPixel(upperLeftColumn);
					if (upperLeftColumn == cells.getMaxDataColumn()) {
						sumColumnsWidthCoveredByWatermark = watermarkBoundingBoxWidth;
					}
				}
				sumRowHeightCoveredByWatermark = 0;
			}
			sumRowHeightCoveredByWatermark += cells.getRowHeightPixel(upperLeftRow);
		}
	}

}
