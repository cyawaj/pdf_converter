package me.season.pdfconvert.convert;

import java.io.File;

import javax.annotation.PostConstruct;

import org.jodconverter.DocumentConverter;
import org.jodconverter.office.OfficeException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import me.season.pdfconvert.service.ConverterManager;

/**
 * 使用JODConverter将word/excel/ppt转PDF
 * 
 * @xavier
 * @date 2019-03-19
 */
@Component
public class OfficeConverter implements Converter {
	@Autowired
	private DocumentConverter converter;

	@PostConstruct
	public void init() {
		ConverterManager.register("doc", this);
		ConverterManager.register("docx", this);
		ConverterManager.register("docm", this);
		ConverterManager.register("xls", this);
		ConverterManager.register("xlsx", this);
		ConverterManager.register("ppt", this);
		ConverterManager.register("pptx", this);
	}

	@Override
	public boolean convert(String source, String target) {
		try {
			converter.convert(new File(source)).to(new File(target)).execute();
			return true;
		} catch (OfficeException e) {
		}
		return false;
	}

}
