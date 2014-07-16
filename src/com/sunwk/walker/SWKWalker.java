package com.sunwk.walker;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.file.FileVisitOption;
import java.nio.file.FileVisitResult;
import java.nio.file.FileVisitor;
import java.nio.file.Files;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.attribute.FileTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.EnumSet;
import java.util.Iterator;
import java.util.StringTokenizer;

import org.apache.lucene.analysis.Analyzer;
import org.apache.lucene.analysis.standard.StandardAnalyzer;
import org.apache.lucene.document.Document;
import org.apache.lucene.document.Field;
import org.apache.lucene.document.LongField;
import org.apache.lucene.document.StringField;
import org.apache.lucene.document.TextField;
import org.apache.lucene.index.IndexWriter;
import org.apache.lucene.index.IndexWriterConfig;
import org.apache.lucene.index.IndexWriterConfig.OpenMode;
import org.apache.lucene.index.Term;
import org.apache.lucene.store.Directory;
import org.apache.lucene.store.FSDirectory;
import org.apache.lucene.util.Version;
import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Notes;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.record.TextHeaderAtom;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class SWKWalker implements FileVisitor{
	ArrayList<String> wordsarray = new ArrayList<>();
	ArrayList<String> documents = new ArrayList<>();
	
	IndexWriter writer = null;
	
	public SWKWalker(String words){
		wordsarray.clear();
		documents.clear();
		
		StringTokenizer st = new StringTokenizer(words, ",");
		
		while(st.hasMoreElements()){
			wordsarray.add(st.nextToken().trim());
		}
	}
	
	void search(Path file) throws IOException{
		String name = file.getFileName().toString();
		//String full_name = file.normalize().toString();
		int mid = name.lastIndexOf(".");
		String ext = name.substring(mid + 1, name.length());
		
		System.out.println("Path name:" + file.toString());
		System.out.println("current file:" + name);
		
		// 확장명이 pdf인 경우
		try{
			Document doc = new Document();
			Field pathField = new StringField("path", file.toString(), Field.Store.YES);
			doc.add(pathField);
			
			FileTime lastModifiedTime = (FileTime)Files.getAttribute(file, "basic:lastModifiedTime", LinkOption.NOFOLLOW_LINKS);
			
			doc.add(new LongField("modified", lastModifiedTime.toMillis(), Field.Store.NO));
			
			if(ext.equalsIgnoreCase("pdf")){			
				searchInPDF_PDFBox(file.toString(), doc);			
			}
			
			if(ext.equalsIgnoreCase("doc")||ext.equalsIgnoreCase("docx")){
				searchInWord(file.toString(), doc);
			}
			
			if(ext.equalsIgnoreCase("ppt")){
				searchInPPT(file.toString(), doc);
			}
			
			if(ext.equalsIgnoreCase("xls")){
				searchInExcel(file.toString(), doc);
			}
			
			if((ext.equalsIgnoreCase("txt")) || (ext.equalsIgnoreCase("xml")
												  || ext.equalsIgnoreCase("html"))
												  || ext.equalsIgnoreCase("htm") || ext.equalsIgnoreCase("xhtml") || ext.equalsIgnoreCase("rtf")){
				searchInText(file, doc);
				
			}
			
			if (writer.getConfig().getOpenMode() == OpenMode.CREATE) {
		        // New index, so we just add the document (no old document can be there):
		        System.out.println("adding " + file);
		        
					writer.addDocument(doc);
				
		      } else {
		        System.out.println("updating " + file);
		        writer.updateDocument(new Term("path", file.toString()), doc);
		      }
			
		}catch(IOException e){
			e.printStackTrace();
		}
	}
	
	boolean searchInText(Path file, Document lucene_doc){
		boolean flag = false;
		Charset charset  = Charset.forName("UTF-8");
		 
		try(BufferedReader reader = Files.newBufferedReader(file, charset)){
			String line = null;
			
			OUTERMOST:
			while((line = reader.readLine())!=null){
				lucene_doc.add(new TextField("contents", line, Field.Store.YES ));
			}
		}catch(IOException e){
			
		}finally{
			return flag;
		}
	}

	void searchInExcel(String file, Document lucene_doc){
		Row row;
		Cell cell;
		String text;

		InputStream xls = null;
		
		try{
			xls = new FileInputStream(file);
			HSSFWorkbook wb = new HSSFWorkbook(xls);
			
			int sheets = wb.getNumberOfSheets();
			
			OUTERMOST:
			for(int i=0; i < sheets; i++){
				HSSFSheet sheet = wb.getSheetAt(i);
				
				Iterator<Row> row_iterator = sheet.rowIterator();
				while(row_iterator.hasNext()){
					row = (Row)row_iterator.next();
					Iterator<Cell> cell_iterator = row.cellIterator();
					while(cell_iterator.hasNext()){
						cell=cell_iterator.next();
						
						int type = cell.getCellType();
						if(type == HSSFCell.CELL_TYPE_STRING){
							text = cell.getStringCellValue();
							
							lucene_doc.add(new TextField("contens", text, Field.Store.YES ));

						}
					}
				}
			}
		}catch(IOException e){			
		}finally{
			try{
				if(xls!=null){
					xls.close();
				}
			}catch(IOException e){
					
			}

		}
	}
	
	void searchInPPT(String file, Document lucene_doc){
//		boolean flag = false;
		InputStream fis = null;
		String text;
		
		try{
			fis = new FileInputStream(new File(file));
			POIFSFileSystem fs = new POIFSFileSystem(fis);
			HSLFSlideShow show = new HSLFSlideShow(fs);
			
			SlideShow ss = new SlideShow(show);
			Slide[] slides = ss.getSlides();
			
			OUTERMOST:
			for(int i=0; i < slides.length; i++){
				TextRun[] runs = slides[i].getTextRuns();
				for(int j = 0; j < runs.length; j++){
					TextRun run = runs[j];
					if(run.getRunType() == TextHeaderAtom.TITLE_TYPE){
						text = run.getText();
					}else{
						text = run.getRunType() + " " + run.getText();
					}
					
					lucene_doc.add(new TextField("contens", text, Field.Store.YES ));

				}
				
				Notes notes = slides[i].getNotesSheet();				
				if(notes != null){
					runs = notes.getTextRuns();
					for(int j = 0; j < runs.length; j++){
						text = runs[j].getText();
						lucene_doc.add(new TextField("contens", text, Field.Store.YES ));

					}
				}
			}
			
		}catch(IOException e){
			
		}finally{
			try{
				if(fis!=null){
					fis.close();
				}
			}catch(IOException e){
				
			}			
		}
	}
	
	boolean searchInWord(String file, Document lucene_doc){
		boolean flag = false;
		
		POIFSFileSystem fs = null;
		try{
			fs = new POIFSFileSystem(new FileInputStream(file));
			
			HWPFDocument doc = new HWPFDocument(fs);
			WordExtractor we = new WordExtractor(doc);
			String[] paragraphs = we.getParagraphText();
			
			for(String paragraph : paragraphs)
				lucene_doc.add(new TextField("contens", paragraph, Field.Store.YES ));
			

		}catch(Exception e){
			
		}finally{
			return flag;
		}
	}
	
	void searchInPDF_PDFBox(String file, Document doc){
		PDFParser parser = null;
		String parsedText = null;
		PDFTextStripper pdfStripper = null;
		PDDocument pdDoc = null;
		COSDocument cosDoc = null;
		boolean flag = false;
		int page = 0;
		
		File pdf = new File(file);
		
		try{
			parser = new PDFParser(new FileInputStream(pdf));
			parser.parse();
			
			cosDoc = parser.getDocument();
			pdfStripper = new PDFTextStripper();
			pdDoc = new PDDocument(cosDoc);
			
			OUTERMOST:
			while(page < pdDoc.getNumberOfPages()){
				page++;
				pdfStripper.setStartPage(page);
				pdfStripper.setEndPage(page + 1);
				parsedText = pdfStripper.getText(pdDoc);
				

				doc.add(new TextField("contens", parsedText, Field.Store.YES ));
			}
		}catch(Exception e){
			
		}finally{
			try{
				if(cosDoc!=null){
					cosDoc.close();
				}
				if(pdDoc!=null){
					pdDoc.close();
				}
			}catch(Exception e){
				
			}
		}
	}

	@Override
	public FileVisitResult preVisitDirectory(Object dir,
			BasicFileAttributes attrs) throws IOException {
		// TODO Auto-generated method stub
				
		return FileVisitResult.CONTINUE;
	}

	@Override
	public FileVisitResult visitFile(Object file, BasicFileAttributes attrs)
			throws IOException {
		// TODO Auto-generated method stub
		search((Path) file);
				
		return FileVisitResult.CONTINUE;
	}

	@Override
	public FileVisitResult visitFileFailed(Object file, IOException exc)
			throws IOException {
		// TODO Auto-generated method stub
		return FileVisitResult.CONTINUE;
	}

	@Override
	public FileVisitResult postVisitDirectory(Object dir, IOException exc)
			throws IOException {
		// TODO Auto-generated method stub
		
		System.out.println("Visited: " + (Path) dir);
		
		return FileVisitResult.CONTINUE;
	}
	
	public static void main(String[] args) throws IOException{
		String usage = "java SWKWalker"
                + " [-index INDEX_PATH] [-docs DOCS_PATH] [-update]\n\n"
                + "This indexes the documents in DOCS_PATH, creating a Lucene index"
                + "in INDEX_PATH that can be searched with SearchFiles";
		String indexPath = "index";
		String docsPath = null;
		boolean create = true;
		for(int i=0;i<args.length;i++) {
			if ("-index".equals(args[i])) {
				indexPath = args[i+1];
				i++;
			} else if ("-docs".equals(args[i])) {
				docsPath = args[i+1];
				i++;
			} else if ("-update".equals(args[i])) {
				create = false;
			}
		}

		if (docsPath == null) {
			System.err.println("Usage: " + usage);
			System.exit(1);
		}

		final Path docPath = Paths.get(docsPath);
		
		if(docPath == null){
			System.out.println("Document directory '" +docsPath+ "' does not exist or is not readable, please check the path");
			System.exit(1);
		}
		
		Directory dir = FSDirectory.open(new File(indexPath));
		Analyzer analyzer = new StandardAnalyzer(Version.LUCENE_4_9);
		IndexWriterConfig iwc = new IndexWriterConfig(Version.LUCENE_4_9, analyzer);
		
		if (create) {
	        // Create a new index in the directory, removing any
	        // previously indexed documents:
			iwc.setOpenMode(OpenMode.CREATE);
		} else {
			// Add new documents to an existing index:
			iwc.setOpenMode(OpenMode.CREATE_OR_APPEND);
		}
		
		String words = "test, words";
		SWKWalker walk= new SWKWalker(words);		
		walk.writer = new IndexWriter(dir, iwc);
		
		EnumSet<FileVisitOption> opts = EnumSet.of(FileVisitOption.FOLLOW_LINKS);
		
		Date start = new Date();
		Files.walkFileTree(docPath, opts, Integer.MAX_VALUE, walk);
		
		walk.writer.close();
		Date end = new Date();
		
		System.out.println("-------------------------------------------------------------------");
		for(String path_string : walk.documents){
			System.out.println(path_string);			
		}
		System.out.println("-------------------------------------------------------------------");
		System.out.println(end.getTime() - start.getTime() + " total milliseconds");
	}

}
