package me.season.pdfconvert.util;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.*;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


/**
 * PDF工具类
 */
public class PDFUtils {
	/**
	 * 合并PDF
	 * 
	 * @param target
	 * @param sources
	 */
	public static void mergePDF(String target, String... sources) {
		mergeFiles(target, sources, true);
	}

	public static void mergeFiles(String result, String[] files, boolean smart) {
		Document document = new Document();
		PdfCopy copy;
		try {
			if (smart)
				copy = new PdfSmartCopy(document, new FileOutputStream(result));
			else
				copy = new PdfCopy(document, new FileOutputStream(result));
			document.open();
			PdfReader[] reader = new PdfReader[3];
			for (int i = 0; i < files.length; i++) {
				reader[i] = new PdfReader(files[i]);
				copy.addDocument(reader[i]);
				copy.freeReader(reader[i]);
				reader[i].close();
			}
		} catch (Exception e) {

		} finally {
			document.close();
		}

	}

	/**
	 * 修改Excel文档页面设置-sheet设置为一页宽，多页高
	 * 
	 * @param file
	 * @return
	 */
	public static boolean modifyPageSetting(File file) {
		if (file != null && file.exists()) {
			String filename = file.getName();
			// 文件后缀
			String suffix = filename.substring(filename.lastIndexOf("."));
			boolean isXLS = suffix.equalsIgnoreCase(".xls");
			boolean isXLSX = suffix.equalsIgnoreCase(".xlsx");
			if (!isXLS && !isXLSX) {
				// 非excel文档不用修改
				return false;
			}
			// 临时文件-原文件的备份
			File tmpFile = new File(file.getParent() + File.separator + "tmp_" + file.getName());
			try {
				FileUtils.copyFile(file, tmpFile);
			} catch (IOException e) {
				return false;
			}
			Workbook workbook = null;
			try (FileInputStream input = new FileInputStream(tmpFile);
					FileOutputStream output = new FileOutputStream(file)) {
				workbook = isXLS ? new HSSFWorkbook(input) : new XSSFWorkbook(input);
				int numberOfSheets = workbook.getNumberOfSheets();
				for (int i = 0; i < numberOfSheets; i++) {
					Sheet sheet = workbook.getSheetAt(i);
					sheet.setFitToPage(true);
					sheet.setAutobreaks(true);
					PrintSetup printSetup = sheet.getPrintSetup();
					printSetup.setFitWidth((short) 1);
					printSetup.setFitHeight((short) 0);
				}
				workbook.write(output);
				return true;
			} catch (Exception e) {
			} finally {
				if (workbook != null) {
					try {
						workbook.close();
					} catch (IOException ex) {
					}
				}
				// 删除临时文件
				tmpFile.delete();
			}
		}
		return false;
	}
}
